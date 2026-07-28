"""Robust parsing/normalization layer for 1C Excel exports."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

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
    "warehouse": ("склад", "место хранения"),
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


def map_input_columns(columns: List[str], aliases: Dict[str, Sequence[str]]) -> Dict[str, str]:
    """Map canonical keys to detected physical column names."""
    mapping: Dict[str, str] = {}
    for canonical, names in aliases.items():
        for col in columns:
            if not col:
                continue
            for alias in names:
                a = alias.casefold()
                if a == col or a in col or col in a:
                    mapping[canonical] = col
                    break
            if canonical in mapping:
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


def build_canonical_stock_df(
    raw: pd.DataFrame, columns: List[str], mapping: Dict[str, str]
) -> pd.DataFrame:
    """Build canonical stock DataFrame: sku, name, stock, uom, warehouse, sku_key."""
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
    uom_s = col_series("uom").map(safe_str) if "uom" in mapping else pd.Series([""] * len(raw))
    warehouse_s = (
        col_series("warehouse").map(safe_str) if "warehouse" in mapping else pd.Series([""] * len(raw))
    )
    group_s = col_series("group").map(safe_str) if "group" in mapping else pd.Series([""] * len(raw))

    out = pd.DataFrame()
    out["sku"] = [
        _choose_sku(article, code, barcode, name)
        for article, code, barcode, name in zip(article_s, code_s, barcode_s, name_s)
    ]
    out["name"] = name_s
    out["stock"] = qty_s
    out["uom"] = uom_s
    out["warehouse"] = warehouse_s
    out["group"] = group_s
    out["sku_key"] = out["sku"].map(normalize_text)

    out = out.dropna(how="all")
    out = out[out["sku_key"] != ""].copy()
    out = out[~out["name"].str.casefold().str.contains("итог", na=False)].copy()
    out = out[~((out["stock"] == 0) & (out["name"].str.strip() == ""))].copy()

    agg = (
        out.groupby("sku_key", as_index=False)
        .agg(
            sku=("sku", "first"),
            name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
            stock=("stock", "sum"),
            uom=("uom", "first"),
            warehouse=("warehouse", lambda s: ", ".join(sorted({x for x in s if safe_str(x)}))),
        )
        .reset_index(drop=True)
    )
    return agg


def build_canonical_sales_df(
    raw: pd.DataFrame, columns: List[str], mapping: Dict[str, str], header_row: int
) -> pd.DataFrame:
    """Build canonical sales DataFrame: sku, name, date, qty, amount, sku_key."""
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

    qty_label = mapping.get("qty", "")
    amount_label = mapping.get("amount", "")
    qty_cols: Dict[int, pd.Timestamp] = {}
    amount_cols: Dict[Tuple[int, pd.Timestamp], int] = {}

    # Date row is usually one line above main headers in 1C wide sales report.
    date_row = header_row - 1 if header_row > 0 else header_row
    date_candidates = raw.iloc[date_row].tolist() if len(raw) else []

    for idx, col_name in enumerate(columns):
        if idx >= len(date_candidates):
            continue
        parsed_date = pd.to_datetime(date_candidates[idx], errors="coerce", dayfirst=True)
        if pd.isna(parsed_date):
            continue
        col_is_qty = qty_label and (col_name == qty_label or qty_label in col_name)
        col_is_amount = amount_label and (col_name == amount_label or amount_label in col_name)
        if col_is_qty:
            qty_cols[idx] = pd.Timestamp(parsed_date).normalize()
        elif col_is_amount:
            amount_cols[(idx, pd.Timestamp(parsed_date).normalize())] = idx

    if not qty_cols:
        # Fallback for "long" files with explicit date column + qty.
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
            }
        )
        out["sku_key"] = out["sku"].map(normalize_text)
        out = out.dropna(subset=["date"])
        out = out[out["sku_key"] != ""].copy()
        out = out[~out["name"].str.casefold().str.contains("итог", na=False)].copy()
        return out.reset_index(drop=True)

    records: List[Dict[str, object]] = []
    data_start = header_row + 1
    for r in range(data_start, len(raw)):
        name = safe_str(name_s.iloc[r])
        code = safe_str(code_s.iloc[r])
        article = safe_str(article_s.iloc[r])
        barcode = safe_str(barcode_s.iloc[r])
        group = safe_str(group_s.iloc[r])

        sku = _choose_sku(article, code, barcode, name)
        if not sku:
            continue
        if "итог" in name.casefold():
            continue
        if not (code or article or barcode or group):
            # Skips group/subtotal lines from 1C export.
            continue

        for qty_col, day in qty_cols.items():
            qty = safe_float(raw.iat[r, qty_col], 0.0)
            if qty == 0.0:
                continue
            amount = 0.0
            for (amount_col, amount_day), _ in amount_cols.items():
                if amount_day == day:
                    amount = safe_float(raw.iat[r, amount_col], 0.0)
                    break
            records.append(
                {
                    "sku": sku,
                    "name": name,
                    "date": day,
                    "qty": qty,
                    "amount": amount,
                }
            )

    out = pd.DataFrame.from_records(records)
    if out.empty:
        raise ParserError("После разбора файл продаж не содержит валидных строк.")
    out["sku_key"] = out["sku"].map(normalize_text)
    out = out[out["sku_key"] != ""].copy()
    return out.reset_index(drop=True)


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
