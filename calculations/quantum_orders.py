"""Навешивание кванта и округление заказов/перемещений."""

from __future__ import annotations

import math

import pandas as pd

from config.quantum_library import apply_quantum_ceil, lookup_quantum
from data.store_utils import is_central_warehouse


def attach_quantum_column(df: pd.DataFrame) -> pd.DataFrame:
    """Добавляет колонку quantum (≥1) по библиотеке / разбору имени."""
    out = df.copy()
    values = []
    for _, row in out.iterrows():
        values.append(
            int(
                lookup_quantum(
                    sku_key=row.get("sku_key", ""),
                    name=row.get("name", ""),
                    sku=row.get("sku", ""),
                    article=row.get("article", ""),
                    code=row.get("code", ""),
                    barcode=row.get("barcode", ""),
                )
            )
        )
    out["quantum"] = values
    return out


def round_orders_to_quantum(df: pd.DataFrame, qty_col: str = "recommended_order") -> pd.DataFrame:
    """Округляет положительный заказ вверх до кванта; 0 остаётся 0."""
    out = df.copy()
    if qty_col not in out.columns:
        return out
    if "quantum" not in out.columns:
        out = attach_quantum_column(out)
    rounded = []
    for _, row in out.iterrows():
        qty = float(row.get(qty_col, 0) or 0)
        if qty <= 0:
            rounded.append(0.0)
            continue
        q = int(row.get("quantum", 1) or 1)
        rounded.append(apply_quantum_ceil(qty, q))
    out[qty_col] = rounded
    out["order_need_flag"] = out[qty_col] > 0
    return out


def quantum_floor_qty(qty: float, quantum: int) -> float:
    """Максимум, кратный кванту и не больше qty (для перемещений со склада)."""
    q = float(qty or 0.0)
    pack = int(quantum or 0)
    if q <= 0 or pack <= 1:
        return max(0.0, q) if pack <= 1 else 0.0
    n = int(math.floor(q / pack + 1e-12))
    return float(n * pack)


def apply_central_supplier_order_from_stores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Заказ поставщику на «Склад основной1» = сумма потребности всех магазинов (после перемещений).

    На розничных строках recommended_order остаётся (матрица / точки).
    На строке ЦС recommended_order / supplier_order_qty = сумма по SKU (с округлением до кванта).
    Если строки ЦС нет — supplier_order_qty ставится на первую розничную строку SKU
    (лист 09 подпишет магазин как Склад основной1).
    """
    out = df.copy()
    if out.empty or "store" not in out.columns or "sku_key" not in out.columns:
        out["supplier_order_qty"] = out.get("recommended_order", 0.0)
        return out

    if "quantum" not in out.columns:
        out = attach_quantum_column(out)

    retail_mask = ~out["store"].map(is_central_warehouse)
    central_mask = out["store"].map(is_central_warehouse)
    out["supplier_order_qty"] = 0.0

    sums = (
        out.loc[retail_mask]
        .groupby("sku_key", dropna=False)["recommended_order"]
        .sum()
        .astype(float)
        .to_dict()
    )

    for sku_key, total in sums.items():
        total = float(total or 0.0)
        if total <= 0:
            continue
        sample = out.index[out["sku_key"] == sku_key]
        quantum = int(out.at[sample[0], "quantum"] or 1) if len(sample) else 1
        order_qty = apply_quantum_ceil(total, quantum)
        c_idx = out.index[central_mask & (out["sku_key"] == sku_key)].tolist()
        if c_idx:
            primary = c_idx[0]
            out.at[primary, "recommended_order"] = order_qty
            out.at[primary, "supplier_order_qty"] = order_qty
            for extra in c_idx[1:]:
                out.at[extra, "recommended_order"] = 0.0
                out.at[extra, "supplier_order_qty"] = 0.0
        else:
            r_idx = out.index[retail_mask & (out["sku_key"] == sku_key)].tolist()
            if r_idx:
                out.at[r_idx[0], "supplier_order_qty"] = order_qty

    # ЦС без розничной потребности — не заказываем отдельно
    for idx in out.index[central_mask]:
        sku = out.at[idx, "sku_key"]
        if float(sums.get(sku, 0) or 0) <= 0:
            out.at[idx, "recommended_order"] = 0.0
            out.at[idx, "supplier_order_qty"] = 0.0

    out["order_need_flag"] = out["recommended_order"] > 0
    return out
