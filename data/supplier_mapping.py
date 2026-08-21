"""
Постоянный слой привязки SKU → поставщик/контрагент.

Источник: data/reference/Привязка_SKU_к_контрагенту.xlsx
Выбор поставщика в приложении — опциональный.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd

from utils.helpers import normalize_text, safe_str

logger = logging.getLogger("zakupki_forecast.supplier_mapping")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_MAPPING_PATH = (
    PROJECT_ROOT / "data" / "reference" / "Привязка_SKU_к_контрагенту.xlsx"
)

SUPPLIER_NONE_LABEL = "Не выбран / общий расчёт"

COLUMN_ALIASES = {
    "name": ("номенклатура", "наименование", "товар"),
    "supplier": ("поставщик", "контрагент", "поставщик/контрагент"),
    "article": ("артикул", "sku"),
    "code": ("код", "код номенклатуры"),
    "barcode": ("штрихкод", "штрих-код"),
}


@dataclass
class SupplierMappingResult:
    """Каноническая таблица привязки и служебные пометки."""

    frame: pd.DataFrame
    ambiguous_articles: List[Dict[str, object]] = field(default_factory=list)
    source_path: str = ""
    sheet_name: str = ""
    header_row: int = 0

    @property
    def suppliers(self) -> List[str]:
        if self.frame.empty or "supplier_name" not in self.frame.columns:
            return []
        return sorted(
            {
                safe_str(x)
                for x in self.frame["supplier_name"].tolist()
                if safe_str(x)
            }
        )


class SupplierMappingError(ValueError):
    """Ошибка чтения/нормализации файла привязки."""


def _normalize_header(value: object) -> str:
    text = safe_str(value).replace("\n", " ")
    text = " ".join(text.split()).strip().casefold()
    return "" if not text or text.startswith("unnamed") else text


def _detect_header_row(raw: pd.DataFrame) -> int:
    limit = min(len(raw), 40)
    best_row, best_score = -1, -1
    for r in range(limit):
        headers = [_normalize_header(v) for v in raw.iloc[r].tolist()]
        joined = " | ".join(h for h in headers if h)
        score = 0
        for aliases in COLUMN_ALIASES.values():
            if any(a in joined for a in aliases):
                score += 1
        if "поставщик" in joined:
            score += 3
        if score > best_score:
            best_score = score
            best_row = r
    if best_row < 0 or best_score < 3:
        raise SupplierMappingError(
            "Не удалось найти строку заголовка в файле привязки SKU к поставщику."
        )
    return best_row


def _map_columns(headers: List[str]) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for idx, header in enumerate(headers):
        if not header:
            continue
        for canonical, aliases in COLUMN_ALIASES.items():
            if canonical in mapping:
                continue
            if any(a == header or a in header for a in aliases):
                mapping[canonical] = idx
                break
    return mapping


def _pick_data_start(raw: pd.DataFrame, header_row: int, col_map: Dict[str, int]) -> int:
    """Первая строка данных после заголовка (пропуск служебных подзаголовков)."""
    name_idx = col_map.get("name")
    supplier_idx = col_map.get("supplier")
    code_idx = col_map.get("code")
    article_idx = col_map.get("article")
    for r in range(header_row + 1, len(raw)):
        name = safe_str(raw.iat[r, name_idx]) if name_idx is not None else ""
        supplier = safe_str(raw.iat[r, supplier_idx]) if supplier_idx is not None else ""
        code = safe_str(raw.iat[r, code_idx]) if code_idx is not None else ""
        article = safe_str(raw.iat[r, article_idx]) if article_idx is not None else ""
        # После ffill поставщика достаточно номенклатуры/кода.
        if (supplier or code or article) and (code or article or name):
            return r
    return header_row + 1


def load_supplier_mapping(path: str | Path | None = None) -> SupplierMappingResult:
    """Читает и нормализует файл привязки SKU → поставщик."""
    path = Path(path) if path else DEFAULT_MAPPING_PATH
    if not path.exists():
        raise SupplierMappingError(f"Файл привязки не найден: {path}")
    if path.stat().st_size == 0:
        raise SupplierMappingError(f"Файл привязки пустой: {path}")

    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception as exc:
        raise SupplierMappingError(
            f"Не удалось открыть файл привязки «{path.name}»: {exc}"
        ) from exc

    sheet = xl.sheet_names[0]
    raw = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=object)
    if raw is None or raw.empty:
        raise SupplierMappingError("Файл привязки не содержит данных.")

    header_row = _detect_header_row(raw)
    headers = [_normalize_header(v) for v in raw.iloc[header_row].tolist()]
    col_map = _map_columns(headers)
    if "supplier" not in col_map:
        raise SupplierMappingError("В файле привязки не найдена колонка «Поставщик».")
    if "code" not in col_map and "article" not in col_map:
        raise SupplierMappingError(
            "В файле привязки не найдены колонки «Код» / «Артикул»."
        )

    data_start = _pick_data_start(raw, header_row, col_map)
    records: List[Dict[str, object]] = []
    # Forward-fill поставщика: в выгрузке 1С значение часто только в merged-ячейке.
    last_supplier = ""
    for r in range(data_start, len(raw)):
        supplier_raw = safe_str(raw.iat[r, col_map["supplier"]]).strip()
        if supplier_raw:
            last_supplier = supplier_raw
        supplier = last_supplier
        name = safe_str(raw.iat[r, col_map["name"]]).strip() if "name" in col_map else ""
        code = safe_str(raw.iat[r, col_map["code"]]).strip() if "code" in col_map else ""
        article = (
            safe_str(raw.iat[r, col_map["article"]]).strip() if "article" in col_map else ""
        )
        barcode = (
            safe_str(raw.iat[r, col_map["barcode"]]).strip() if "barcode" in col_map else ""
        )
        if not supplier:
            continue
        if not (code or article or barcode or name):
            continue
        # Строки-заголовки групп без идентификатора товара пропускаем.
        if not (code or article or barcode):
            continue
        records.append(
            {
                "supplier_name": supplier,
                "sku": code or article or barcode,
                "item_name": name,
                "code": code,
                "article": article,
                "barcode": barcode,
                "code_key": normalize_text(code) if code else "",
                "article_key": normalize_text(article) if article else "",
                "barcode_key": normalize_text(barcode) if barcode else "",
            }
        )

    frame = pd.DataFrame.from_records(records)
    if frame.empty:
        raise SupplierMappingError(
            "После нормализации файл привязки не содержит валидных строк "
            "(нужны поставщик и код/артикул)."
        )

    # Неоднозначные артикулы: один артикул → несколько поставщиков.
    ambiguous: List[Dict[str, object]] = []
    if "article_key" in frame.columns:
        grouped = (
            frame[frame["article_key"] != ""]
            .groupby("article_key")["supplier_name"]
            .nunique()
        )
        amb_keys = set(grouped[grouped > 1].index.tolist())
        if amb_keys:
            for key in sorted(amb_keys):
                suppliers = sorted(
                    frame.loc[frame["article_key"] == key, "supplier_name"].unique().tolist()
                )
                ambiguous.append({"article_key": key, "suppliers": suppliers})
            # Исключаем неоднозначные артикулы из ключей сопоставления.
            frame.loc[frame["article_key"].isin(amb_keys), "article_key"] = ""
            logger.warning(
                "Найдено %s неоднозначных артикулов (исключены из сопоставления)",
                len(amb_keys),
            )

    # Дедуп по поставщик+код+артикул
    frame = frame.drop_duplicates(
        subset=["supplier_name", "code_key", "article_key", "barcode_key"],
        keep="first",
    ).reset_index(drop=True)

    result = SupplierMappingResult(
        frame=frame,
        ambiguous_articles=ambiguous,
        source_path=str(path),
        sheet_name=sheet,
        header_row=header_row + 1,
    )
    logger.info(
        "Привязка SKU→поставщик: %s строк, %s поставщиков, ambiguous=%s",
        len(frame),
        len(result.suppliers),
        len(ambiguous),
    )
    return result


@lru_cache(maxsize=1)
def get_cached_supplier_mapping() -> SupplierMappingResult:
    """Кэш для UI/Streamlit: файл привязки читается один раз за процесс."""
    return load_supplier_mapping()


def clear_supplier_mapping_cache() -> None:
    """Сброс кэша (для тестов / обновления справочника)."""
    get_cached_supplier_mapping.cache_clear()


def list_suppliers(mapping: Optional[SupplierMappingResult] = None) -> List[str]:
    """Полный список поставщиков из справочника (после ffill)."""
    mapping = mapping or get_cached_supplier_mapping()
    return mapping.suppliers


def sku_keys_for_supplier(
    supplier_name: str,
    mapping: Optional[SupplierMappingResult] = None,
) -> Set[str]:
    """Набор sku_key для выбранного поставщика (code + article + barcode)."""
    mapping = mapping or get_cached_supplier_mapping()
    supplier_name = safe_str(supplier_name).strip()
    if not supplier_name or supplier_name == SUPPLIER_NONE_LABEL:
        return set()

    part = mapping.frame[mapping.frame["supplier_name"] == supplier_name]
    keys: Set[str] = set()
    for col in ("code_key", "article_key", "barcode_key"):
        if col not in part.columns:
            continue
        keys.update(k for k in part[col].tolist() if k)
    return keys


def _supplier_catalog_rows(
    supplier_name: str,
    mapping: SupplierMappingResult,
) -> pd.DataFrame:
    """Ассортимент поставщика из справочника в каноническом формате остатков."""
    part = mapping.frame[mapping.frame["supplier_name"] == supplier_name].copy()
    if part.empty:
        return pd.DataFrame(columns=["sku", "name", "stock", "uom", "warehouse", "sku_key"])

    rows: List[Dict[str, object]] = []
    for _, row in part.iterrows():
        sku = safe_str(row.get("sku"))
        code_key = safe_str(row.get("code_key"))
        article_key = safe_str(row.get("article_key"))
        barcode_key = safe_str(row.get("barcode_key"))
        # Как в loaders: article → code → barcode → name
        sku_key = article_key or code_key or barcode_key or normalize_text(sku)
        if not sku_key:
            continue
        rows.append(
            {
                "sku": sku,
                "name": safe_str(row.get("item_name")),
                "stock": 0.0,
                "uom": "",
                "warehouse": "",
                "sku_key": sku_key,
                "alt_keys": {k for k in (sku_key, code_key, article_key, barcode_key) if k},
            }
        )

    if not rows:
        return pd.DataFrame(columns=["sku", "name", "stock", "uom", "warehouse", "sku_key"])

    # Схлопываем дубли по основному sku_key.
    merged: Dict[str, Dict[str, object]] = {}
    for row in rows:
        key = str(row["sku_key"])
        if key not in merged:
            merged[key] = row
            continue
        prev = merged[key]
        prev_name = safe_str(prev.get("name"))
        new_name = safe_str(row.get("name"))
        if len(new_name) > len(prev_name):
            prev["name"] = new_name
        prev_alts = set(prev.get("alt_keys") or set())
        prev_alts |= set(row.get("alt_keys") or set())
        prev["alt_keys"] = prev_alts

    out = pd.DataFrame(
        [
            {
                "sku": v["sku"],
                "name": v["name"],
                "stock": 0.0,
                "uom": "",
                "warehouse": "",
                "sku_key": v["sku_key"],
                "alt_keys": v.get("alt_keys") or {v["sku_key"]},
            }
            for v in merged.values()
        ]
    )
    return out.reset_index(drop=True)


def _empty_sales_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=["sku", "name", "date", "qty", "amount", "sku_key"])


def filter_frames_by_supplier(
    stock: pd.DataFrame,
    sales: pd.DataFrame,
    supplier_name: Optional[str],
    mapping: Optional[SupplierMappingResult] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, object]]:
    """
    Опциональная фильтрация входных таблиц по поставщику.

    Если поставщик не выбран — возвращает исходные frame без изменений.

    Если поставщик выбран:
    - ассортимент строится по справочнику привязки (даже без строк в остатках);
    - фактические остатки из файла накладываются поверх (нет в файле → 0);
    - продажи фильтруются по SKU поставщика (могут быть пустыми).
    """
    supplier_name = safe_str(supplier_name).strip()
    info: Dict[str, object] = {
        "supplier_selected": False,
        "supplier_name": "",
        "sku_keys": 0,
        "stock_rows": len(stock),
        "sales_rows": len(sales),
        "allow_empty_sales": False,
    }
    if not supplier_name or supplier_name == SUPPLIER_NONE_LABEL:
        return stock, sales, info

    mapping = mapping or get_cached_supplier_mapping()
    keys = sku_keys_for_supplier(supplier_name, mapping)
    if not keys:
        raise SupplierMappingError(
            f"Для поставщика «{supplier_name}» не найдено ни одного SKU в файле привязки."
        )

    catalog = _supplier_catalog_rows(supplier_name, mapping)
    if catalog.empty:
        raise SupplierMappingError(
            f"Для поставщика «{supplier_name}» справочник привязки пуст."
        )

    # Фактические остатки по ключам поставщика.
    stock_in = stock[stock["sku_key"].isin(keys)].copy() if not stock.empty else stock
    sales_f = sales[sales["sku_key"].isin(keys)].copy() if not sales.empty else _empty_sales_frame()

    # База = весь ассортимент поставщика (stock=0), поверх — остатки из файла.
    stock_f = catalog.copy()
    if stock_in is not None and not stock_in.empty:
        stock_by_key = {
            safe_str(r["sku_key"]): r
            for _, r in stock_in.iterrows()
            if safe_str(r.get("sku_key"))
        }
        for idx, row in stock_f.iterrows():
            alt_keys = set(row.get("alt_keys") or set()) | {safe_str(row.get("sku_key"))}
            matched = None
            for ak in alt_keys:
                if ak in stock_by_key:
                    matched = stock_by_key[ak]
                    break
            if matched is None:
                continue
            stock_f.at[idx, "stock"] = float(matched.get("stock", 0) or 0)
            src_name = safe_str(matched.get("name"))
            if src_name and len(src_name) >= len(safe_str(row.get("name"))):
                stock_f.at[idx, "name"] = src_name
            src_sku = safe_str(matched.get("sku"))
            if src_sku:
                stock_f.at[idx, "sku"] = src_sku
            stock_f.at[idx, "uom"] = safe_str(matched.get("uom"))
            stock_f.at[idx, "warehouse"] = safe_str(matched.get("warehouse"))

    # В расчётный pipeline не передаём служебную колонку alt_keys.
    if "alt_keys" in stock_f.columns:
        stock_f = stock_f.drop(columns=["alt_keys"])
    info.update(
        {
            "supplier_selected": True,
            "supplier_name": supplier_name,
            "sku_keys": len(keys),
            "stock_rows": len(stock_f),
            "sales_rows": len(sales_f),
            "stock_before": len(stock),
            "sales_before": len(sales),
            "stock_from_file": 0 if stock_in is None else len(stock_in),
            "allow_empty_sales": bool(sales_f.empty),
        }
    )

    logger.info(
        "Фильтр поставщика «%s»: catalog=%s, stock_file=%s→%s, sales %s→%s, keys=%s",
        supplier_name,
        len(catalog),
        info["stock_from_file"],
        info["stock_rows"],
        info["sales_before"],
        info["sales_rows"],
        info["sku_keys"],
    )
    return stock_f, sales_f, info
