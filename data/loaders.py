"""
Загрузка Excel-файлов остатков и продаж из 1С.

Важно: наименования читаются как строки целиком (dtype=str для текстовых полей),
без усечения и без автоматического преобразования в числа.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config.column_mapping import (
    REQUIRED_SALES_FIELDS,
    REQUIRED_STOCK_FIELDS,
    SALES_COLUMN_ALIASES,
    STOCK_COLUMN_ALIASES,
)
from utils.helpers import normalize_text, safe_float, safe_str

logger = logging.getLogger("zakupki_forecast.loaders")


class ColumnMappingError(Exception):
    """Не удалось сопоставить обязательные колонки."""


def _pick_sheet_name(xl: "pd.ExcelFile") -> str | int:
    """
    Выбирает лист для загрузки из уже открытого файла.
    Приоритет: «Данные» (шаблон) → первый лист.
    Лист «Инструкция» никогда не читается как таблица данных.
    """
    names = list(xl.sheet_names)
    for preferred in ("Данные", "Данные ", "Data", "data"):
        for name in names:
            if safe_str(name).strip().casefold() == preferred.strip().casefold():
                return name
    # Первый лист, который не инструкция
    for name in names:
        low = safe_str(name).strip().casefold()
        if low not in {"инструкция", "instruction", "readme", "справка"}:
            return name
    return 0


# Порядок движков чтения Excel:
#   openpyxl  — основной для .xlsx;
#   calamine  — устойчивый резерв для нестандартных выгрузок из 1С
#               (в т.ч. с перепутанными путями/регистром внутри архива,
#               а также .xls / .xlsb / .ods);
#   xlrd      — старые .xls, если calamine недоступен.
_EXCEL_ENGINES: Tuple[Optional[str], ...] = ("openpyxl", "calamine", "xlrd")


def _open_excel_file(path: Path) -> Tuple[Optional[str], "pd.ExcelFile"]:
    """
    Открывает Excel, перебирая движки по очереди.
    Возвращает (engine, ExcelFile). Если ни один движок не справился —
    бросает понятную ошибку (частый случай — «битая»/нестандартная выгрузка 1С).
    """
    errors: List[str] = []
    for engine in _EXCEL_ENGINES:
        try:
            xl = pd.ExcelFile(path, engine=engine)
            # Форсируем разбор структуры: ленивые движки падают именно здесь
            _ = xl.sheet_names
            logger.debug("Файл %s открыт движком %s", path.name, engine)
            return engine, xl
        except Exception as exc:  # перебираем движки, накапливаем причины
            errors.append(f"{engine or 'auto'}: {exc}")

    detail = "; ".join(errors)
    raise ValueError(
        f"Не удалось прочитать Excel-файл «{path.name}». "
        "Похоже, файл повреждён или сохранён в нестандартном формате "
        "(частая проблема выгрузок из 1С). Откройте его в Excel/LibreOffice "
        "и пересохраните как «Книга Excel (*.xlsx)», затем загрузите снова.\n"
        f"Технические детали: {detail}"
    )


def _read_excel_raw(path: Path) -> pd.DataFrame:
    """Читает Excel с защитой от пустых/битых файлов; перебирает движки."""
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")
    if path.stat().st_size == 0:
        raise ValueError(f"Файл пустой: {path}")

    # dtype=object — не теряем ведущие нули в артикулах и длинные названия.
    # Читаем лист «Данные», чтобы лист «Инструкция» в шаблоне не ломал загрузку.
    _engine, xl = _open_excel_file(path)
    sheet = _pick_sheet_name(xl)
    df = pd.read_excel(xl, sheet_name=sheet, dtype=object)

    if df is None or df.empty:
        raise ValueError(
            f"В файле нет данных: {path.name}. "
            "Заполните лист «Данные» (первая строка — заголовки колонок)."
        )

    # Удаляем полностью пустые строки
    df = df.dropna(how="all").copy()
    if df.empty:
        raise ValueError(f"После очистки файл пуст: {path.name}")

    # Нормализуем имена колонок: trim, без потери кириллицы
    df.columns = [safe_str(c).strip() for c in df.columns]
    return df


def _resolve_columns(
    df: pd.DataFrame,
    aliases: Dict[str, List[str]],
    required: Tuple[str, ...],
) -> Dict[str, str]:
    """
    Сопоставляет канонические поля с реальными колонками файла.
    Возвращает dict: canonical -> actual_column_name.
    """
    columns_lower = {c.casefold(): c for c in df.columns}
    mapping: Dict[str, str] = {}

    for canonical, names in aliases.items():
        found: Optional[str] = None
        for alias in names:
            key = alias.casefold()
            if key in columns_lower:
                found = columns_lower[key]
                break
            # Частичное совпадение (например «Количество остаток»)
            for col_key, col_name in columns_lower.items():
                if key in col_key or col_key in key:
                    found = col_name
                    break
            if found:
                break
        if found:
            mapping[canonical] = found

    missing = [f for f in required if f not in mapping]
    if missing:
        available = ", ".join(df.columns.astype(str))
        raise ColumnMappingError(
            "Не найдены обязательные колонки: "
            + ", ".join(missing)
            + f".\nДоступные столбцы: {available}\n"
            "Добавьте синонимы в config/column_mapping.py"
        )
    return mapping


def load_stock_file(path: str | Path) -> pd.DataFrame:
    """
    Загружает файл остатков.
    Возвращает DataFrame с колонками: sku, name, stock, uom, warehouse, sku_key.
    """
    path = Path(path)
    logger.info("Загрузка остатков: %s", path.name)
    raw = _read_excel_raw(path)
    mapping = _resolve_columns(raw, STOCK_COLUMN_ALIASES, REQUIRED_STOCK_FIELDS)

    out = pd.DataFrame()
    out["sku"] = raw[mapping["sku"]].map(safe_str)
    out["name"] = raw[mapping["name"]].map(safe_str)  # полное наименование без обрезки
    out["stock"] = raw[mapping["stock"]].map(lambda x: safe_float(x, 0.0))

    if "uom" in mapping:
        out["uom"] = raw[mapping["uom"]].map(safe_str)
    else:
        out["uom"] = ""

    if "warehouse" in mapping:
        out["warehouse"] = raw[mapping["warehouse"]].map(safe_str)
    else:
        out["warehouse"] = ""

    out["sku_key"] = out["sku"].map(normalize_text)

    # Агрегация дублей по артикулу (сумма остатков), имя — первое непустое длинное
    out = (
        out.groupby("sku_key", as_index=False)
        .agg(
            sku=("sku", "first"),
            name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
            stock=("stock", "sum"),
            uom=("uom", "first"),
            warehouse=("warehouse", lambda s: ", ".join(sorted({x for x in s if x}))),
        )
    )
    out = out[out["sku_key"] != ""].copy()
    logger.info("Остатки: %s позиций", len(out))
    return out.reset_index(drop=True)


def load_sales_file(path: str | Path) -> pd.DataFrame:
    """
    Загружает файл продаж.
    Возвращает DataFrame: sku, name, date, qty, amount, sku_key.
    """
    path = Path(path)
    logger.info("Загрузка продаж: %s", path.name)
    raw = _read_excel_raw(path)
    mapping = _resolve_columns(raw, SALES_COLUMN_ALIASES, REQUIRED_SALES_FIELDS)

    out = pd.DataFrame()
    out["sku"] = raw[mapping["sku"]].map(safe_str)
    out["name"] = raw[mapping["name"]].map(safe_str)
    out["date"] = pd.to_datetime(raw[mapping["date"]], errors="coerce", dayfirst=True)
    out["qty"] = raw[mapping["qty"]].map(lambda x: safe_float(x, 0.0))

    if "amount" in mapping:
        out["amount"] = raw[mapping["amount"]].map(lambda x: safe_float(x, 0.0))
    else:
        out["amount"] = 0.0

    out["sku_key"] = out["sku"].map(normalize_text)
    out = out.dropna(subset=["date"])
    out = out[out["sku_key"] != ""].copy()

    if out.empty:
        raise ValueError("В файле продаж нет строк с корректной датой.")

    logger.info(
        "Продажи: %s строк, период %s — %s",
        len(out),
        out["date"].min().date(),
        out["date"].max().date(),
    )
    return out.reset_index(drop=True)


def detect_sales_date_range(sales: pd.DataFrame) -> Tuple[pd.Timestamp, pd.Timestamp]:
    """Возвращает min/max дату продаж для подсказки периода."""
    return sales["date"].min(), sales["date"].max()
