"""
Слой привязки SKU → поставщик/контрагент (релиз 3).

Источник по умолчанию ищется в нескольких местах (см. resolve_mapping_path).
Чтение устойчиво к «шумным» выгрузкам 1С: служебные строки, заголовок не в первой
строке, разные названия колонок.

Выбор поставщика в UI — ОПЦИОНАЛЕН. Если поставщик не выбран, этот модуль
не влияет на общий расчёт.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from config.settings import PROJECT_ROOT, SETTINGS
from utils.helpers import normalize_text, safe_str

logger = logging.getLogger("zakupki_forecast.supplier_mapping")

# Канонические поля → возможные названия колонок (без выдумывания жёсткой схемы файла)
SUPPLIER_MAPPING_ALIASES: Dict[str, Sequence[str]] = {
    "supplier_name": (
        "контрагент",
        "поставщик",
        "наименование контрагента",
        "контрагент.наименование",
        "поставщик.наименование",
        "название контрагента",
        "название поставщика",
        "организация",
        "partner",
        "supplier",
        "vendor",
    ),
    "supplier_id": (
        "код контрагента",
        "код поставщика",
        "контрагент.код",
        "инн",
        "инн контрагента",
        "id контрагента",
        "id поставщика",
        "supplier id",
        "supplier_id",
    ),
    "sku": (
        "артикул",
        "код",
        "код номенклатуры",
        "номенклатура.код",
        "sku",
        "код товара",
        "штрихкод",
        "штрих-код",
    ),
    "item_name": (
        "номенклатура",
        "наименование",
        "наименование номенклатуры",
        "номенклатура.наименование",
        "товар",
        "название",
        "полное наименование",
        "item",
        "product",
    ),
}

CANONICAL_COLUMNS = ("supplier_name", "supplier_id", "sku", "item_name", "sku_key")


@dataclass
class SupplierMappingResult:
    """Результат загрузки и нормализации привязки."""

    mapping: pd.DataFrame
    source_path: Optional[Path] = None
    sheet_name: str = ""
    header_row: int = -1
    column_mapping: Dict[str, str] = field(default_factory=dict)
    ambiguous_skus: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    loaded: bool = False

    @property
    def suppliers(self) -> List[str]:
        if self.mapping.empty or "supplier_name" not in self.mapping.columns:
            return []
        return sorted(
            {
                safe_str(x)
                for x in self.mapping["supplier_name"].tolist()
                if safe_str(x)
            },
            key=lambda s: s.casefold(),
        )


def resolve_mapping_path(explicit: str | Path | None = None) -> Optional[Path]:
    """
    Порядок поиска файла привязки:
    1) явный путь (аргумент / настройка);
    2) встроенный data/Привязка_SKU_к_контрагенту.xlsx;
    3) исходный Windows-путь пользователя (если доступен в среде).
    """
    candidates: List[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    cfg = SETTINGS.get("supplier_mapping_path")
    if cfg:
        candidates.append(Path(cfg))
    candidates.append(PROJECT_ROOT / "data" / "Привязка_SKU_к_контрагенту.xlsx")
    # Исходный путь с рабочей станции (может быть недоступен в Linux/Cloud)
    candidates.append(
        Path(
            r"C:\Users\Администратор\Documents\МР\новый период работы\Закупка"
            r"\Привязка_SKU_к_контрагенту.xlsx"
        )
    )
    candidates.append(
        Path(
            "/mnt/c/Users/Администратор/Documents/МР/новый период работы/Закупка"
            "/Привязка_SKU_к_контрагенту.xlsx"
        )
    )

    seen = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def _normalize_header_cell(value: object) -> str:
    text = safe_str(value).replace("\n", " ")
    text = " ".join(text.split()).strip().casefold()
    if not text or text.startswith("unnamed"):
        return ""
    return text


def _score_header_row(values: Iterable[object]) -> int:
    norms = [_normalize_header_cell(v) for v in values]
    norms = [n for n in norms if n]
    if not norms:
        return 0
    score = 0
    joined = " | ".join(norms)
    for aliases in SUPPLIER_MAPPING_ALIASES.values():
        for alias in aliases:
            if any(alias == n or alias in n for n in norms):
                score += 3
                break
        else:
            if any(a in joined for a in aliases):
                score += 1
    return score


def _detect_header_row(raw: pd.DataFrame, max_scan: int = 40) -> int:
    best_row, best_score = 0, -1
    limit = min(len(raw), max_scan)
    for i in range(limit):
        score = _score_header_row(raw.iloc[i].tolist())
        if score > best_score:
            best_score = score
            best_row = i
    return best_row if best_score > 0 else 0


def _map_columns(headers: Sequence[str]) -> Dict[str, str]:
    """canonical_field -> original header text."""
    mapping: Dict[str, str] = {}
    used = set()
    for field_name, aliases in SUPPLIER_MAPPING_ALIASES.items():
        for idx, header in enumerate(headers):
            if idx in used or not header:
                continue
            if any(alias == header or alias in header for alias in aliases):
                mapping[field_name] = header
                used.add(idx)
                break
    return mapping


def _pick_best_sheet(xl: pd.ExcelFile) -> Tuple[str, pd.DataFrame, int, Dict[str, str]]:
    best: Optional[Tuple[int, str, pd.DataFrame, int, Dict[str, str]]] = None
    for sheet in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=object)
        if raw is None or raw.empty:
            continue
        raw = raw.dropna(axis=1, how="all")
        header_row = _detect_header_row(raw)
        headers = [_normalize_header_cell(v) for v in raw.iloc[header_row].tolist()]
        colmap = _map_columns(headers)
        score = len(colmap) * 10
        if "supplier_name" in colmap:
            score += 20
        if "sku" in colmap:
            score += 20
        if "item_name" in colmap:
            score += 5
        candidate = (score, sheet, raw, header_row, colmap)
        if best is None or score > best[0]:
            best = candidate
    if best is None:
        raise ValueError("В файле привязки нет читаемых листов.")
    return best[1], best[2], best[3], best[4]


def _empty_mapping() -> pd.DataFrame:
    return pd.DataFrame(columns=list(CANONICAL_COLUMNS))


def normalize_supplier_mapping(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Канонический вид: supplier_name, supplier_id, sku, item_name, sku_key.
    Дубли и неоднозначные привязки обрабатываются явно.
    """
    warnings: List[str] = []
    if df.empty:
        return _empty_mapping(), [], warnings

    out = df.copy()
    for col in ("supplier_name", "supplier_id", "sku", "item_name"):
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].map(lambda v: safe_str(v).strip())

    out = out[(out["sku"] != "") & (out["supplier_name"] != "")].copy()
    if out.empty:
        warnings.append("После очистки не осталось строк с заполненными SKU и поставщиком.")
        return _empty_mapping(), [], warnings

    out["sku_key"] = out["sku"].map(normalize_text)
    out = out[out["sku_key"] != ""].copy()

    # Уникальные пары sku_key + supplier
    out = out.drop_duplicates(subset=["sku_key", "supplier_name"], keep="first")

    # Неоднозначные SKU (несколько поставщиков)
    counts = out.groupby("sku_key")["supplier_name"].nunique()
    ambiguous_keys = counts[counts > 1].index.tolist()
    ambiguous_skus: List[str] = []
    if ambiguous_keys:
        amb_rows = out[out["sku_key"].isin(ambiguous_keys)]
        ambiguous_skus = sorted({safe_str(s) for s in amb_rows["sku"].tolist() if safe_str(s)})
        warnings.append(
            f"Найдено {len(ambiguous_skus)} SKU с привязкой к нескольким поставщикам. "
            "Для фильтрации используется первое соответствие (стабильная сортировка)."
        )
        # Стабильный выбор: сортировка по поставщику, оставляем первое
        out = out.sort_values(["sku_key", "supplier_name"], kind="mergesort")
        out = out.drop_duplicates(subset=["sku_key"], keep="first")

    out = out.reset_index(drop=True)
    return out[list(CANONICAL_COLUMNS)], ambiguous_skus, warnings


def load_supplier_mapping(path: str | Path | None = None) -> SupplierMappingResult:
    """Загружает и нормализует файл привязки SKU к контрагенту."""
    resolved = resolve_mapping_path(path)
    if resolved is None:
        msg = (
            "Файл привязки SKU к контрагенту не найден. "
            "Общий расчёт доступен; режим «по поставщику» недоступен, "
            "пока файл не положен в data/Привязка_SKU_к_контрагенту.xlsx."
        )
        logger.warning(msg)
        return SupplierMappingResult(
            mapping=_empty_mapping(),
            loaded=False,
            warnings=[msg],
        )

    try:
        xl = pd.ExcelFile(resolved, engine="openpyxl")
        sheet_name, raw, header_row, colmap = _pick_best_sheet(xl)
    except Exception as exc:
        msg = f"Не удалось прочитать файл привязки «{resolved}»: {exc}"
        logger.error(msg)
        return SupplierMappingResult(
            mapping=_empty_mapping(),
            source_path=resolved,
            loaded=False,
            warnings=[msg],
        )

    if "supplier_name" not in colmap or "sku" not in colmap:
        msg = (
            f"В файле «{resolved.name}» (лист «{sheet_name}») не найдены колонки "
            "поставщика и/или SKU. Проверьте заголовки."
        )
        logger.error(msg)
        return SupplierMappingResult(
            mapping=_empty_mapping(),
            source_path=resolved,
            sheet_name=sheet_name,
            header_row=header_row,
            column_mapping=colmap,
            loaded=False,
            warnings=[msg],
        )

    headers_raw = [safe_str(v).strip() for v in raw.iloc[header_row].tolist()]
    headers_norm = [_normalize_header_cell(v) for v in headers_raw]
    body = raw.iloc[header_row + 1 :].copy()
    body.columns = [
        headers_raw[i] if i < len(headers_raw) and headers_raw[i] else f"col_{i}"
        for i in range(len(body.columns))
    ]
    body = body.dropna(how="all")

    # Построить DataFrame по каноническим полям через нормализованные заголовки
    series_map: Dict[str, pd.Series] = {}
    for field_name, norm_header in colmap.items():
        # найти исходный индекс колонки
        try:
            idx = headers_norm.index(norm_header)
        except ValueError:
            continue
        col_name = body.columns[idx]
        series_map[field_name] = body[col_name]

    framed = pd.DataFrame(series_map)
    mapping, ambiguous, warnings = normalize_supplier_mapping(framed)
    result = SupplierMappingResult(
        mapping=mapping,
        source_path=resolved,
        sheet_name=sheet_name,
        header_row=header_row,
        column_mapping=colmap,
        ambiguous_skus=ambiguous,
        warnings=warnings,
        loaded=not mapping.empty,
    )
    logger.info(
        "Привязка SKU→поставщик: path=%s sheet=%s rows=%s suppliers=%s ambiguous=%s",
        resolved,
        sheet_name,
        len(mapping),
        len(result.suppliers),
        len(ambiguous),
    )
    return result


def get_skus_for_supplier(mapping: pd.DataFrame, supplier_name: str) -> set[str]:
    """Множество sku_key для выбранного поставщика."""
    if mapping is None or mapping.empty or not supplier_name:
        return set()
    name = safe_str(supplier_name).strip()
    if not name:
        return set()
    mask = mapping["supplier_name"].map(lambda v: safe_str(v).strip()) == name
    keys = mapping.loc[mask, "sku_key"].tolist()
    return {k for k in keys if k}


def filter_frames_by_supplier(
    stock: pd.DataFrame,
    sales: pd.DataFrame,
    mapping: pd.DataFrame,
    supplier_name: str,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, object]]:
    """
    Фильтрует остатки и продажи по SKU выбранного поставщика.
    Не меняет формулы — только отбор строк до расчёта.
    """
    sku_keys = get_skus_for_supplier(mapping, supplier_name)
    info: Dict[str, object] = {
        "supplier_name": safe_str(supplier_name),
        "mapping_sku_count": len(sku_keys),
        "stock_before": len(stock),
        "sales_before": len(sales),
    }
    if not sku_keys:
        info["stock_after"] = 0
        info["sales_after"] = 0
        info["warning"] = (
            f"Для поставщика «{supplier_name}» в привязке нет SKU "
            "или ни один SKU не совпал с ключами в остатках/продажах."
        )
        return stock.iloc[0:0].copy(), sales.iloc[0:0].copy(), info

    stock_f = stock[stock["sku_key"].isin(sku_keys)].copy()
    sales_f = sales[sales["sku_key"].isin(sku_keys)].copy()
    info["stock_after"] = len(stock_f)
    info["sales_after"] = len(sales_f)
    info["matched_stock_skus"] = int(stock_f["sku_key"].nunique()) if not stock_f.empty else 0
    info["matched_sales_skus"] = int(sales_f["sku_key"].nunique()) if not sales_f.empty else 0
    return stock_f, sales_f, info
