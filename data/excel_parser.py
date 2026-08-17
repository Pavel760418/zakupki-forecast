"""Robust parsing/normalization layer for 1C Excel exports."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from data.store_utils import canon_store_name, store_key
from utils.helpers import normalize_text, safe_float, safe_str


@dataclass
class ParseDiagnostics:
    file_type: str
    selected_sheet: str
    header_row: int
    recognized_columns: List[str]
    mapping: Dict[str, str]
    join_keys_priority: List[str]

    def as_log_message(self) -> str:
        return (
            f"[{self.file_type}] sheet='{self.selected_sheet}', "
            f"header_row={self.header_row}, "
            f"recognized={self.recognized_columns}, "
            f"mapping={self.mapping}, "
            f"join_keys_priority={self.join_keys_priority}"
        )


class ParserError(ValueError):
    """Known parser/normalization error with user-friendly text."""


EXCEL_ENGINES: Tuple[Optional[str], ...] = ("openpyxl", "calamine", "xlrd")

# Primary candidates from new 1C templates
STOCK_ALIASES: Dict[str, Sequence[str]] = {
    "name": ("номенклатура", "наименование", "товар"),
    "code": ("код", "код номенклатуры"),
    "article": ("артикул", "sku"),
    "barcode": ("штрих-код", "штрихкод"),
    "group": ("входит в группу", "группа"),
    "qty": ("количество", "остаток", "в наличии"),
    "amount": ("сумма", "стоимость"),
    "uom": ("единица", "ед. изм", "ед."),
    "warehouse": ("склад", "место хранения", "магазин", "подразделение", "торговая точка"),
}

SALES_ALIASES: Dict[str, Sequence[str]] = {
    "name": ("номенклатура", "наименование", "товар"),
    "code": ("код", "код номенклатуры"),
    "article": ("артикул", "sku"),
    "barcode": ("штрих-код", "штрихкод"),
    "group": ("входит в группу", "группа"),
    "qty": ("количество", "продажи", "продано", "кол-во"),
    "amount": ("сумма", "выручка", "стоимость"),
    "date": ("дата", "период", "день"),
    "uom": ("единица", "ед. изм", "ед."),
    "warehouse": ("склад", "место хранения", "магазин", "подразделение", "торговая точка"),
}


def normalize_columns(columns: Iterable[object]) -> List[str]:
    """Normalize column captions from noisy 1C Excel exports."""
    normalized: List[str] = []
    for col in columns:
        text = safe_str(col).replace("\n", " ")
        text = " ".join(text.split())
        text = text.strip().casefold()
        if not text or text.startswith("unnamed"):
            normalized.append("")
        else:
            normalized.append(text)
    return normalized


def _open_excel(path: Path) -> Tuple[Optional[str], pd.ExcelFile]:
    errors: List[str] = []
    for engine in EXCEL_ENGINES:
        try:
            xl = pd.ExcelFile(path, engine=engine)
            _ = xl.sheet_names
            return engine, xl
        except Exception as exc:
            errors.append(f"{engine or 'auto'}: {exc}")
    raise ParserError(
        f"Не удалось открыть файл «{path.name}». "
        "Откройте его в Excel и пересохраните как *.xlsx. "
        f"Технические детали: {'; '.join(errors)}"
    )


def _read_sheet_raw(xl: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=object)
    if raw is None or raw.empty:
        return pd.DataFrame()
    # Keep original row indexes for accurate diagnostics/header positions.
    raw = raw.dropna(axis=1, how="all")
    return raw


def _sheet_score(raw: pd.DataFrame, aliases: Dict[str, Sequence[str]]) -> int:
    if raw.empty:
        return -10_000
    limit = min(len(raw), 80)
    score = 0
    for r in range(limit):
        row_norm = normalize_columns(raw.iloc[r].tolist())
        joined = " | ".join([x for x in row_norm if x])
        for names in aliases.values():
            for alias in names:
                a = alias.casefold()
                if a in joined:
                    score += 3
        # Prefer "header-like" short text rows
        non_empty = [x for x in row_norm if x]
        if 3 <= len(non_empty) <= 25:
            score += 1
    return score


def detect_sheet(xl: pd.ExcelFile, aliases: Dict[str, Sequence[str]]) -> str:
    """Detect best worksheet with actual table, not service text."""
    names = list(xl.sheet_names)
    scored: List[Tuple[int, str]] = []
    for name in names:
        raw = _read_sheet_raw(xl, name)
        score = _sheet_score(raw, aliases)
        low = safe_str(name).strip().casefold()
        if low in {"инструкция", "instruction", "readme", "справка"}:
            score -= 100
        scored.append((score, name))
    scored.sort(reverse=True)
    if not scored:
        raise ParserError("В файле не найдено листов.")
    return scored[0][1]


def detect_header_row(raw: pd.DataFrame, aliases: Dict[str, Sequence[str]]) -> int:
    """Detect row index (0-based) with main column headers."""
    if raw.empty:
        raise ParserError("Лист пустой — не удалось найти заголовок таблицы.")
    limit = min(len(raw), 120)
    best_row = -1
    best_score = -10_000
    for r in range(limit):
        norm = normalize_columns(raw.iloc[r].tolist())
        non_empty = [v for v in norm if v]
        if not non_empty:
            continue
        row_text = " | ".join(non_empty)
        alias_hits = 0
        for names in aliases.values():
            if any(alias.casefold() in row_text for alias in names):
                alias_hits += 1
        score = alias_hits * 10 + min(len(non_empty), 15)
        if score > best_score:
            best_score = score
            best_row = r
    if best_row < 0 or best_score < 15:
        raise ParserError(
            "Не удалось определить строку заголовка таблицы. "
            "Проверьте, что в файле есть колонки номенклатуры/кода/количества."
        )
    return best_row


def _compose_headers(raw: pd.DataFrame, header_row: int) -> List[str]:
    """Compose robust headers using row above for merged group captions."""
    primary = raw.iloc[header_row].tolist()
    primary_norm = normalize_columns(primary)
    if header_row == 0:
        return primary_norm

    upper = raw.iloc[header_row - 1].tolist()
    upper_norm = normalize_columns(upper)
    upper_ffill: List[str] = []
    last = ""
    for value in upper_norm:
        if value:
            last = value
        upper_ffill.append(last)

    headers: List[str] = []
    for up, cur in zip(upper_ffill, primary_norm):
        if cur:
            headers.append(cur)
        elif up:
            headers.append(up)
        else:
            headers.append("")
    return headers


def _alias_hits_column(col: str, alias: str) -> bool:
    """Match alias to a header. Do not use 'col in alias' — «код» ложно цеплялся к «штрихкод»."""
    if not col or not alias:
        return False
    a = alias.casefold()
    c = col.casefold()
    return c == a or a in c


def map_input_columns(columns: List[str], aliases: Dict[str, Sequence[str]]) -> Dict[str, str]:
    """Map canonical keys to detected physical column names."""
    mapping: Dict[str, str] = {}
    # Штрихкод раньше кода — более длинный алиас не должен проигрывать «код».
    key_order = [k for k in aliases if k == "barcode"] + [k for k in aliases if k != "barcode"]
    used_cols: set[str] = set()
    for canonical in key_order:
        names = aliases[canonical]
        for col in columns:
            if not col or col in used_cols:
                continue
            if any(_alias_hits_column(col, alias) for alias in names):
                mapping[canonical] = col
                used_cols.add(col)
                break
    return mapping


def _column_positions(columns: List[str]) -> Dict[str, List[int]]:
    pos: Dict[str, List[int]] = {}
    for idx, col in enumerate(columns):
        pos.setdefault(col, []).append(idx)
    return pos


def _choose_sku(*values: object) -> str:
    for value in values:
        text = safe_str(value).strip()
        if text:
            return text
    return ""


def _is_section_header(name: str, code: str, article: str, barcode: str) -> bool:
    """Строка-группировка 1С: «Адлер» / «Обособленное подразделение Флагман» без кода товара."""
    if not name:
        return False
    low = name.casefold()
    if "итог" in low or low.startswith("*"):
        return True
    return not (code or article or barcode)


def _ffill_dates(values: Sequence[object]) -> List[Optional[pd.Timestamp]]:
    filled: List[Optional[pd.Timestamp]] = []
    last: Optional[pd.Timestamp] = None
    for value in values:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if pd.notna(parsed):
            last = pd.Timestamp(parsed).normalize()
        filled.append(last)
    return filled


def build_canonical_stock_df(
    raw: pd.DataFrame, columns: List[str], mapping: Dict[str, str]
) -> pd.DataFrame:
    """Канон остатков: sku, name, stock, uom, store, barcode, sku_key (строка = SKU × магазин)."""
    positions = _column_positions(columns)

    def col_series(key: str) -> pd.Series:
        if key not in mapping:
            return pd.Series([None] * len(raw))
        idx = positions[mapping[key]][0]
        return raw.iloc[:, idx]

    name_s = col_series("name").map(safe_str)
    code_s = col_series("code").map(safe_str)
    article_s = col_series("article").map(safe_str)
    barcode_s = col_series("barcode").map(safe_str)
    qty_s = col_series("qty").map(lambda x: safe_float(x, 0.0))
    amount_s = (
        col_series("amount").map(lambda x: safe_float(x, 0.0))
        if "amount" in mapping
        else pd.Series([0.0] * len(raw))
    )
    uom_s = col_series("uom").map(safe_str) if "uom" in mapping else pd.Series([""] * len(raw))
    warehouse_s = (
        col_series("warehouse").map(safe_str) if "warehouse" in mapping else pd.Series([""] * len(raw))
    )
    group_s = col_series("group").map(safe_str) if "group" in mapping else pd.Series([""] * len(raw))

    current_store = ""
    records: List[Dict[str, object]] = []
    for i in range(len(raw)):
        name = safe_str(name_s.iloc[i]).strip()
        code = safe_str(code_s.iloc[i]).strip()
        article = safe_str(article_s.iloc[i]).strip()
        barcode = safe_str(barcode_s.iloc[i]).strip()
        if _is_section_header(name, code, article, barcode):
            if "итог" in name.casefold() or name.startswith("*"):
                continue
            current_store = canon_store_name(name)
            continue
        sku = _choose_sku(article, code, barcode, name)
        if not sku or "итог" in name.casefold():
            continue
        if name.casefold() in {"номенклатура", "наименование", "артикул", "код", "склад"}:
            continue
        col_store = canon_store_name(warehouse_s.iloc[i]) if i < len(warehouse_s) else ""
        store = col_store or current_store
        qty = float(qty_s.iloc[i] or 0.0)
        if qty == 0.0 and not name:
            continue
        records.append(
            {
                "sku": sku,
                "name": name,
                "code": code,
                "article": article,
                "barcode": barcode,
                "stock": qty,
                "stock_amount": float(amount_s.iloc[i] or 0.0),
                "uom": safe_str(uom_s.iloc[i]) if i < len(uom_s) else "",
                "store": store,
                "group": safe_str(group_s.iloc[i]) if i < len(group_s) else "",
            }
        )

    out = pd.DataFrame.from_records(records)
    if out.empty:
        return out
    out["sku_key"] = out["sku"].map(normalize_text)
    out["store"] = out["store"].map(lambda x: canon_store_name(x) if x else "")
    out["store_key"] = out["store"].map(store_key)
    out["warehouse"] = out["store"]
    out = out[out["sku_key"] != ""].copy()

    agg = (
        out.groupby(["sku_key", "store_key"], as_index=False, dropna=False)
        .agg(
            sku=("sku", "first"),
            name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
            stock=("stock", "sum"),
            stock_amount=("stock_amount", "sum"),
            uom=("uom", "first"),
            store=("store", "first"),
            warehouse=("store", "first"),
            barcode=("barcode", lambda s: next((x for x in s if safe_str(x)), "")),
            article=("article", "first"),
            code=("code", "first"),
            group=("group", "first"),
        )
        .reset_index(drop=True)
    )
    return agg


def build_canonical_sales_df(
    raw: pd.DataFrame, columns: List[str], mapping: Dict[str, str], header_row: int
) -> pd.DataFrame:
    """Канон продаж: sku, name, date, qty, amount, store, barcode, sku_key."""
    positions = _column_positions(columns)

    def pick_column(key: str) -> pd.Series:
        if key not in mapping:
            return pd.Series([None] * len(raw))
        return raw.iloc[:, positions[mapping[key]][0]]

    name_s = pick_column("name").map(safe_str)
    code_s = pick_column("code").map(safe_str)
    article_s = pick_column("article").map(safe_str)
    barcode_s = pick_column("barcode").map(safe_str)
    group_s = pick_column("group").map(safe_str) if "group" in mapping else pd.Series([""] * len(raw))
    warehouse_s = (
        pick_column("warehouse").map(safe_str) if "warehouse" in mapping else pd.Series([""] * len(raw))
    )

    qty_label = mapping.get("qty", "")
    amount_label = mapping.get("amount", "")
    qty_cols: Dict[int, pd.Timestamp] = {}
    amount_by_date: Dict[pd.Timestamp, int] = {}

    date_row = header_row - 1 if header_row > 0 else header_row
    date_candidates = raw.iloc[date_row].tolist() if len(raw) else []
    filled_dates = _ffill_dates(date_candidates)

    for idx, col_name in enumerate(columns):
        if idx >= len(filled_dates) or filled_dates[idx] is None:
            continue
        day = filled_dates[idx]
        col_is_qty = qty_label and (col_name == qty_label or qty_label in col_name)
        col_is_amount = amount_label and (col_name == amount_label or amount_label in col_name)
        if col_is_qty:
            qty_cols[idx] = day
        elif col_is_amount:
            amount_by_date[day] = idx

    def _with_store(frame: pd.DataFrame) -> pd.DataFrame:
        if "store" not in frame.columns:
            frame["store"] = ""
        frame["store"] = frame["store"].map(lambda x: canon_store_name(x) if x else "")
        frame["store_key"] = frame["store"].map(store_key)
        frame["warehouse"] = frame["store"]
        if "barcode" not in frame.columns:
            frame["barcode"] = ""
        return frame

    if not qty_cols:
        if "date" not in mapping or "qty" not in mapping:
            raise ParserError(
                "В файле продаж не найдены колонки дат и количеств. "
                "Проверьте структуру выгрузки 1С."
            )
        date_s = pd.to_datetime(pick_column("date"), errors="coerce", dayfirst=True)
        qty_s = pick_column("qty").map(lambda x: safe_float(x, 0.0))
        amount_s = (
            pick_column("amount").map(lambda x: safe_float(x, 0.0))
            if "amount" in mapping
            else pd.Series([0.0] * len(raw))
        )
        stores: List[str] = []
        current = ""
        for i in range(len(raw)):
            name = safe_str(name_s.iloc[i]).strip()
            code = safe_str(code_s.iloc[i]).strip()
            article = safe_str(article_s.iloc[i]).strip()
            barcode = safe_str(barcode_s.iloc[i]).strip()
            if _is_section_header(name, code, article, barcode):
                if "итог" not in name.casefold() and not name.startswith("*"):
                    current = canon_store_name(name)
                stores.append("")
                continue
            col_store = canon_store_name(warehouse_s.iloc[i]) if i < len(warehouse_s) else ""
            stores.append(col_store or current)
        out = pd.DataFrame(
            {
                "sku": [
                    _choose_sku(article, code, barcode, name)
                    for article, code, barcode, name in zip(article_s, code_s, barcode_s, name_s)
                ],
                "name": name_s,
                "date": date_s,
                "qty": qty_s,
                "amount": amount_s,
                "group": group_s,
                "barcode": barcode_s,
                "store": stores,
            }
        )
        out["sku_key"] = out["sku"].map(normalize_text)
        out = out.dropna(subset=["date"])
        out = out[out["sku_key"] != ""].copy()
        out = out[~out["name"].str.casefold().str.contains("итог", na=False)].copy()
        return _with_store(out.reset_index(drop=True))

    records: List[Dict[str, object]] = []
    data_start = header_row + 1
    current_store = ""
    for r in range(data_start, len(raw)):
        name = safe_str(name_s.iloc[r]).strip()
        code = safe_str(code_s.iloc[r]).strip()
        article = safe_str(article_s.iloc[r]).strip()
        barcode = safe_str(barcode_s.iloc[r]).strip()
        group = safe_str(group_s.iloc[r]).strip() if r < len(group_s) else ""

        if _is_section_header(name, code, article, barcode):
            if "итог" in name.casefold() or name.startswith("*"):
                continue
            current_store = canon_store_name(name)
            continue

        sku = _choose_sku(article, code, barcode, name)
        if not sku or "итог" in name.casefold():
            continue
        if name.casefold() in {"номенклатура", "наименование", "артикул", "код", "склад"}:
            continue
        if not (code or article or barcode or group):
            continue
        col_store = canon_store_name(warehouse_s.iloc[r]) if r < len(warehouse_s) else ""
        store = col_store or current_store

        for qty_col, day in qty_cols.items():
            qty = safe_float(raw.iat[r, qty_col], 0.0)
            if qty == 0.0:
                continue
            amount_col = amount_by_date.get(day)
            amount = safe_float(raw.iat[r, amount_col], 0.0) if amount_col is not None else 0.0
            records.append(
                {
                    "sku": sku,
                    "name": name,
                    "date": day,
                    "qty": qty,
                    "amount": amount,
                    "barcode": barcode,
                    "store": store,
                    "group": group,
                }
            )

    out = pd.DataFrame.from_records(records)
    if out.empty:
        raise ParserError("После разбора файл продаж не содержит валидных строк.")
    out["sku_key"] = out["sku"].map(normalize_text)
    out = out[out["sku_key"] != ""].copy()
    return _with_store(out.reset_index(drop=True))


def parse_stock_file(path: str | Path) -> Tuple[pd.DataFrame, ParseDiagnostics]:
    path = Path(path)
    if not path.exists():
        raise ParserError(f"Файл остатков не найден: {path}")
    if path.stat().st_size == 0:
        raise ParserError("Файл остатков пустой.")

    _engine, xl = _open_excel(path)
    sheet = detect_sheet(xl, STOCK_ALIASES)
    raw = _read_sheet_raw(xl, sheet)
    if raw.empty:
        raise ParserError("На выбранном листе остатков нет данных.")

    header_row = detect_header_row(raw, STOCK_ALIASES)
    columns = _compose_headers(raw, header_row)
    mapping = map_input_columns(columns, STOCK_ALIASES)

    if "name" not in mapping or "qty" not in mapping:
        raise ParserError(
            "Не найдены ключевые колонки в файле остатков: "
            "ожидаются номенклатура и количество."
        )

    canonical = build_canonical_stock_df(raw, columns, mapping)
    if canonical.empty:
        raise ParserError("Файл остатков прочитан, но строки товаров не найдены.")

    diagnostics = ParseDiagnostics(
        file_type="stock",
        selected_sheet=sheet,
        header_row=header_row + 1,
        recognized_columns=[c for c in columns if c],
        mapping=mapping,
        join_keys_priority=["article", "code", "barcode", "name"],
    )
    return canonical, diagnostics


def parse_sales_file(path: str | Path) -> Tuple[pd.DataFrame, ParseDiagnostics]:
    path = Path(path)
    if not path.exists():
        raise ParserError(f"Файл продаж не найден: {path}")
    if path.stat().st_size == 0:
        raise ParserError("Файл продаж пустой.")

    _engine, xl = _open_excel(path)
    sheet = detect_sheet(xl, SALES_ALIASES)
    raw = _read_sheet_raw(xl, sheet)
    if raw.empty:
        raise ParserError("На выбранном листе продаж нет данных.")

    header_row = detect_header_row(raw, SALES_ALIASES)
    columns = _compose_headers(raw, header_row)
    mapping = map_input_columns(columns, SALES_ALIASES)

    if "name" not in mapping:
        raise ParserError("В файле продаж не найдена колонка номенклатуры.")
    if "qty" not in mapping:
        raise ParserError("В файле продаж не найдена колонка количества.")

    canonical = build_canonical_sales_df(raw, columns, mapping, header_row=header_row)
    if canonical.empty:
        raise ParserError("Файл продаж прочитан, но валидные продажи не извлечены.")

    diagnostics = ParseDiagnostics(
        file_type="sales",
        selected_sheet=sheet,
        header_row=header_row + 1,
        recognized_columns=[c for c in columns if c],
        mapping=mapping,
        join_keys_priority=["article", "code", "barcode", "name"],
    )
    return canonical, diagnostics
