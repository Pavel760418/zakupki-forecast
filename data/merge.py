"""Объединение остатков и продаж в единую витрину для расчётов."""

from __future__ import annotations

import logging
from typing import Tuple

import numpy as np
import pandas as pd

from data.store_utils import NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL, canon_store_name, store_key
from utils.helpers import safe_str

logger = logging.getLogger("zakupki_forecast.merge")

GRAIN_NETWORK = "network"
GRAIN_STORE = "store"


def filter_sales_by_period(
    sales: pd.DataFrame,
    date_from: pd.Timestamp,
    date_to: pd.Timestamp,
) -> pd.DataFrame:
    """Отбирает продажи внутри периода включительно."""
    mask = (sales["date"] >= date_from) & (sales["date"] <= date_to)
    filtered = sales.loc[mask].copy()
    logger.info("Продажи в периоде: %s строк", len(filtered))
    return filtered


def _best_name(*series_list) -> str:
    """Выбирает самое полное (длинное) наименование без усечения."""
    candidates = []
    for s in series_list:
        if s is None:
            continue
        for v in (s if isinstance(s, (list, tuple, pd.Series)) else [s]):
            text = safe_str(v)
            if text:
                candidates.append(text)
    if not candidates:
        return ""
    return max(candidates, key=len)


def _ensure_store_cols(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    if "store" not in out.columns:
        if "warehouse" in out.columns:
            out["store"] = out["warehouse"].map(canon_store_name)
        else:
            out["store"] = ""
    else:
        out["store"] = out["store"].map(lambda x: canon_store_name(x) if safe_str(x).strip() else "")
    out["store_key"] = out["store"].map(store_key)
    if "barcode" not in out.columns:
        out["barcode"] = ""
    if "stock_amount" not in out.columns:
        out["stock_amount"] = 0.0
    return out


def _aggregate_stock(stock: pd.DataFrame, grain: str) -> pd.DataFrame:
    keys = ["sku_key"] if grain == GRAIN_NETWORK else ["sku_key", "store_key"]
    grouped = (
        stock.groupby(keys, as_index=False, dropna=False)
        .agg(
            sku=("sku", "first"),
            name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
            stock=("stock", "sum"),
            stock_amount=("stock_amount", "sum"),
            uom=("uom", "first") if "uom" in stock.columns else ("sku", "first"),
            barcode=("barcode", lambda s: next((safe_str(x) for x in s if safe_str(x)), "")),
            store=("store", "first"),
        )
    )
    if grain == GRAIN_NETWORK:
        grouped["store"] = NETWORK_STORE_LABEL
        grouped["store_key"] = store_key(NETWORK_STORE_LABEL)
    else:
        grouped["store"] = grouped["store"].map(
            lambda x: x if safe_str(x).strip() else UNKNOWN_STORE_LABEL
        )
        grouped["store_key"] = grouped["store"].map(store_key)
    return grouped


def build_product_frame(
    stock: pd.DataFrame,
    sales_period: pd.DataFrame,
    date_from: pd.Timestamp,
    date_to: pd.Timestamp,
    grain: str = GRAIN_NETWORK,
) -> Tuple[pd.DataFrame, dict]:
    """
    Продуктовая витрина.

    grain=network — 1 строка на SKU (сводно по сети).
    grain=store — 1 строка на SKU × магазин.
    """
    days = int((date_to - date_from).days) + 1
    mid = date_from + pd.Timedelta(days=days // 2)
    grain = GRAIN_STORE if grain == GRAIN_STORE else GRAIN_NETWORK

    stock = _ensure_store_cols(stock)
    sales_period = _ensure_store_cols(sales_period)
    stock_agg = _aggregate_stock(stock, grain)

    group_keys = ["sku_key"] if grain == GRAIN_NETWORK else ["sku_key", "store_key"]

    if sales_period.empty:
        sales_agg = pd.DataFrame(columns=group_keys + ["sales_qty", "sales_amount", "sales_name", "last_sale_date", "sku"])
    else:
        sales_work = sales_period.copy()
        if grain == GRAIN_NETWORK:
            sales_work["store_key"] = store_key(NETWORK_STORE_LABEL)
            sales_work["store"] = NETWORK_STORE_LABEL
        else:
            sales_work["store"] = sales_work["store"].map(
                lambda x: x if safe_str(x).strip() else UNKNOWN_STORE_LABEL
            )
            sales_work["store_key"] = sales_work["store"].map(store_key)
        sales_agg = (
            sales_work.groupby(group_keys, as_index=False, dropna=False)
            .agg(
                sales_qty=("qty", "sum"),
                sales_amount=("amount", "sum"),
                sales_name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
                last_sale_date=("date", "max"),
                sku=("sku", "first"),
                store=("store", "first"),
            )
        )

    first_half = sales_period[sales_period["date"] < mid]
    second_half = sales_period[sales_period["date"] >= mid]

    def _half_sum(part: pd.DataFrame) -> pd.Series:
        if part.empty:
            return pd.Series(dtype=float)
        work = part.copy()
        if grain == GRAIN_NETWORK:
            return work.groupby("sku_key")["qty"].sum()
        work["store"] = work["store"].map(lambda x: x if safe_str(x).strip() else UNKNOWN_STORE_LABEL)
        work["store_key"] = work["store"].map(store_key)
        return work.groupby(["sku_key", "store_key"])["qty"].sum()

    h1 = _half_sum(first_half)
    h2 = _half_sum(second_half)

    merged = stock_agg.merge(sales_agg, on=group_keys, how="outer", suffixes=("", "_sales"))

    merged["sku"] = merged.apply(
        lambda r: safe_str(r.get("sku")) or safe_str(r.get("sku_sales")),
        axis=1,
    )
    merged["name"] = merged.apply(
        lambda r: _best_name(r.get("name"), r.get("sales_name")),
        axis=1,
    )
    merged["stock"] = merged["stock"].fillna(0.0).astype(float)
    merged["stock_amount"] = merged.get("stock_amount", 0.0)
    if not isinstance(merged["stock_amount"], pd.Series):
        merged["stock_amount"] = 0.0
    merged["stock_amount"] = pd.to_numeric(merged["stock_amount"], errors="coerce").fillna(0.0)
    merged["sales_qty"] = merged["sales_qty"].fillna(0.0).astype(float)
    merged["sales_amount"] = merged["sales_amount"].fillna(0.0).astype(float)
    merged["uom"] = merged.get("uom", pd.Series([""] * len(merged))).fillna("").map(safe_str)
    merged["barcode"] = merged.get("barcode", pd.Series([""] * len(merged))).fillna("").map(safe_str)
    if "store" not in merged.columns:
        merged["store"] = ""
    if "store_sales" in merged.columns:
        merged["store"] = merged.apply(
            lambda r: safe_str(r.get("store")) or safe_str(r.get("store_sales")),
            axis=1,
        )
    merged["store"] = merged["store"].map(
        lambda x: NETWORK_STORE_LABEL
        if grain == GRAIN_NETWORK
        else (safe_str(x).strip() or UNKNOWN_STORE_LABEL)
    )
    merged["store_key"] = merged["store"].map(store_key)
    merged["warehouse"] = merged["store"]

    if grain == GRAIN_NETWORK:
        merged["sales_h1"] = merged["sku_key"].map(h1).fillna(0.0)
        merged["sales_h2"] = merged["sku_key"].map(h2).fillna(0.0)
    else:
        def _map_half(row, series: pd.Series) -> float:
            key = (row["sku_key"], row["store_key"])
            if key in series.index:
                return float(series.loc[key])
            return 0.0

        merged["sales_h1"] = merged.apply(lambda r: _map_half(r, h1), axis=1)
        merged["sales_h2"] = merged.apply(lambda r: _map_half(r, h2), axis=1)

    if "last_sale_date" in merged.columns:
        merged["days_since_sale"] = merged["last_sale_date"].apply(
            lambda d: (date_to - d).days if pd.notna(d) else days
        )
    else:
        merged["days_since_sale"] = days

    merged = merged.sort_values(["sales_qty", "name"], ascending=[False, True]).reset_index(drop=True)
    merged.insert(0, "row_id", np.arange(1, len(merged) + 1))

    stores_in_play = sorted(
        {
            safe_str(x)
            for x in merged["store"].tolist()
            if safe_str(x) and safe_str(x) not in {NETWORK_STORE_LABEL}
        }
    )
    meta = {
        "date_from": date_from,
        "date_to": date_to,
        "period_days": days,
        "mid_date": mid,
        "items_count": len(merged),
        "total_sales_qty": float(merged["sales_qty"].sum()),
        "total_stock": float(merged["stock"].sum()),
        "grain": grain,
        "stores": stores_in_play,
        "store_count": len(stores_in_play),
    }
    logger.info("Витрина: %s строк, grain=%s, магазинов=%s", len(merged), grain, len(stores_in_play))
    return merged, meta
