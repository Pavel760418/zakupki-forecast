"""
Загрузка и нормализация входных Excel-файлов:
- основной прайс-лист ("Исходные данные");
- опциональный файл продаж.

Модуль требует Excel-файл (.xlsx/.xls) для загрузки прайс-листа.
Файл продаж - опциональная вариация, не обязателен.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from za_price_calculator.config import (
    SOURCE_COLUMNS,
    SOURCE_COLUMN_ALIASES,
    SALES_COLUMNS,
    SALES_COLUMN_ALIASES,
)
from za_price_calculator.exceptions import FileLoadError, ValidationError

logger = logging.getLogger(__name__)

_REQUIRED_SOURCE_MIN = {"Наименование", "Закупочная_цена", "Розничная_цена_ЗЯ"}
_REQUIRED_SALES_MIN = {"Штрихкод"}

_PRICE_RE = re.compile(r"^\s*([\d]+(?:[.,]\d+)?)")


def _normalize_header(raw: object) -> str:
    return str(raw).strip().lower().replace("\u00a0", " ") if raw is not None else ""


def _map_columns(df: pd.DataFrame, aliases: dict, canonical: list) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for col in df.columns:
        norm = _normalize_header(col)
        if norm in aliases:
            rename_map[col] = aliases[norm]
        elif col in canonical:
            rename_map[col] = col
    return df.rename(columns=rename_map)


def _clean_price(value: object):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if pd.notna(value) else None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("\u00a0", " ").replace(" ", "")
    match = _PRICE_RE.match(text)
    if not match:
        return None
    try:
        return float(match.group(1).replace(",", "."))
    except ValueError:
        return None


def _clean_barcode(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()


def load_source_file(file_path, sheet_name=0) -> pd.DataFrame:
    """
    Загружает и нормализует основной прайс-лист.

    :param file_path: путь к Excel-файлу (.xlsx/.xls).
    :param sheet_name: имя или индекс листа. По умолчанию - первый лист.
    :return: DataFrame с колонками SOURCE_COLUMNS, числовыми ценами и строковыми штрихкодами.
    :raises FileLoadError: файл не найден / не читается / неверный формат.
    :raises ValidationError: отсутствуют обязательные колонки или нет валидных строк.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileLoadError(f"Файл не найден: {path}")
    if path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        raise FileLoadError(f"Неподдерживаемый формат файла: {path.suffix}. Ожидается .xlsx/.xls/.xlsm")

    try:
        df_raw = pd.read_excel(path, sheet_name=sheet_name, header=0, engine=None)
    except Exception as exc:
        raise FileLoadError(f"Не удалось прочитать Excel-файл '{path.name}': {exc}") from exc

    if isinstance(df_raw, dict):
        df_raw = next(iter(df_raw.values()))

    df = _map_columns(df_raw, SOURCE_COLUMN_ALIASES, SOURCE_COLUMNS)

    missing_required = _REQUIRED_SOURCE_MIN - set(df.columns)
    if missing_required:
        raise ValidationError(
            "В файле отсутствуют обязательные колонки: "
            f"{', '.join(sorted(missing_required))}. "
            "Проверьте структуру шаблона (лист 'Исходные данные')."
        )

    for col in SOURCE_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df = df[SOURCE_COLUMNS].copy()

    df["Наименование"] = df["Наименование"].apply(
        lambda v: str(v).strip() if v is not None and pd.notna(v) else ""
    )
    df["Штрихкод"] = df["Штрихкод"].apply(_clean_barcode)

    price_cols = [c for c in SOURCE_COLUMNS if c not in ("Наименование", "Штрихкод")]
    for col in price_cols:
        df[col] = df[col].apply(_clean_price)

    df = df[df["Наименование"] != ""].reset_index(drop=True)

    if df.empty:
        raise ValidationError(
            "После очистки данных не осталось валидных строк товаров. "
            "Проверьте, что колонка 'Наименование' заполнена."
        )

    n_no_purchase = df["Закупочная_цена"].isna().sum()
    n_no_retail = df["Розничная_цена_ЗЯ"].isna().sum()
    if n_no_purchase:
        logger.warning("%d строк без закупочной цены - расчёты наценки/маржи по ним будут пустыми", n_no_purchase)
    if n_no_retail:
        logger.warning("%d строк без розничной цены ЗЯ - расчёты по ним будут пустыми", n_no_retail)

    logger.info("Загружено %d позиций из '%s'", len(df), path.name)
    return df


def load_sales_file(file_path, sheet_name=0) -> pd.DataFrame:
    """
    Загружает опциональный файл продаж. Не является обязательным для работы калькулятора.

    :param file_path: путь к Excel-файлу с продажами.
    :param sheet_name: имя/индекс листа.
    :return: DataFrame с колонками SALES_COLUMNS (Штрихкод как строка, числа - float).
    :raises FileLoadError: файл не найден/не читается.
    :raises ValidationError: отсутствует обязательная колонка "Штрихкод".
    """
    path = Path(file_path)
    if not path.exists():
        raise FileLoadError(f"Файл продаж не найден: {path}")
    if path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
        raise FileLoadError(f"Неподдерживаемый формат файла продаж: {path.suffix}")

    try:
        df_raw = pd.read_excel(path, sheet_name=sheet_name, header=0, engine=None)
    except Exception as exc:
        raise FileLoadError(f"Не удалось прочитать файл продаж '{path.name}': {exc}") from exc

    if isinstance(df_raw, dict):
        df_raw = next(iter(df_raw.values()))

    df = _map_columns(df_raw, SALES_COLUMN_ALIASES, SALES_COLUMNS)

    missing_required = _REQUIRED_SALES_MIN - set(df.columns)
    if missing_required:
        raise ValidationError(
            f"В файле продаж отсутствует обязательная колонка: {', '.join(missing_required)}"
        )

    for col in SALES_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df = df[SALES_COLUMNS].copy()
    df["Штрихкод"] = df["Штрихкод"].apply(_clean_barcode)
    for col in ("Кол-во_продаж", "Выручка", "Валовая_прибыль"):
        df[col] = df[col].apply(_clean_price)

    df = df[df["Штрихкод"] != ""].reset_index(drop=True)

    logger.info("Загружено %d строк продаж из '%s'", len(df), path.name)
    return df
