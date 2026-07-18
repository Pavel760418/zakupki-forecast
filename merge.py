"""Объединение остатков и продаж в единую витрину для расчётов."""

from __future__ import annotations

import logging
from typing import Tuple

import numpy as np
import pandas as pd

from utils.helpers import safe_str

logger = logging.getLogger("zakupki_forecast.merge")


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


def build_product_frame(
    stock: pd.DataFrame,
    sales_period: pd.DataFrame,
    date_from: pd.Timestamp,
    date_to: pd.Timestamp,
) -> Tuple[pd.DataFrame, dict]:
    """
    Строит продуктовую витрину:
    - продажи за период, 1-я/2-я половины (для тренда);
    - остаток;
    - справочные атрибуты.

    Возвращает (df, meta).
    """
    days = int((date_to - date_from).days) + 1
    mid = date_from + pd.Timedelta(days=days // 2)

    # Агрегация продаж за период
    if sales_period.empty:
        sales_agg = pd.DataFrame(
            columns=["sku_key", "sales_qty", "sales_amount", "sales_name", "last_sale_date"]
        )
    else:
        sales_agg = (
            sales_period.groupby("sku_key", as_index=False)
            .agg(
                sales_qty=("qty", "sum"),
                sales_amount=("amount", "sum"),
                sales_name=("name", lambda s: max((safe_str(x) for x in s), key=len, default="")),
                last_sale_date=("date", "max"),
                sku=("sku", "first"),
            )
        )

    # Половинки периода для тренда
    first_half = sales_period[sales_period["date"] < mid]
    second_half = sales_period[sales_period["date"] >= mid]

    def _half_sum(part: pd.DataFrame) -> pd.Series:
        if part.empty:
            return pd.Series(dtype=float)
        return part.groupby("sku_key")["qty"].sum()

    h1 = _half_sum(first_half)
    h2 = _half_sum(second_half)

    # Outer join: товары с остатком и/или продажами
    stock_keyed = stock.copy()
    merged = stock_keyed.merge(sales_agg, on="sku_key", how="outer", suffixes=("", "_sales"))

    # SKU / имя без потерь
    merged["sku"] = merged.apply(
        lambda r: safe_str(r.get("sku")) or safe_str(r.get("sku_sales")),
        axis=1,
    )
    merged["name"] = merged.apply(
        lambda r: _best_name(r.get("name"), r.get("sales_name")),
        axis=1,
    )
    merged["stock"] = merged["stock"].fillna(0.0).astype(float)
    merged["sales_qty"] = merged["sales_qty"].fillna(0.0).astype(float)
    merged["sales_amount"] = merged["sales_amount"].fillna(0.0).astype(float)
    merged["uom"] = merged.get("uom", pd.Series([""] * len(merged))).fillna("").map(safe_str)
    merged["warehouse"] = (
        merged.get("warehouse", pd.Series([""] * len(merged))).fillna("").map(safe_str)
    )

    merged["sales_h1"] = merged["sku_key"].map(h1).fillna(0.0)
    merged["sales_h2"] = merged["sku_key"].map(h2).fillna(0.0)

    # Дней с последней продажи (в рамках загруженных данных периода)
    if "last_sale_date" in merged.columns:
        merged["days_since_sale"] = merged["last_sale_date"].apply(
            lambda d: (date_to - d).days if pd.notna(d) else days
        )
    else:
        merged["days_since_sale"] = days

    merged = merged.sort_values(["sales_qty", "name"], ascending=[False, True]).reset_index(drop=True)
    merged.insert(0, "row_id", np.arange(1, len(merged) + 1))

    meta = {
        "date_from": date_from,
        "date_to": date_to,
        "period_days": days,
        "mid_date": mid,
        "items_count": len(merged),
        "total_sales_qty": float(merged["sales_qty"].sum()),
        "total_stock": float(merged["stock"].sum()),
    }
    logger.info("Витрина: %s SKU", len(merged))
    return merged, meta
