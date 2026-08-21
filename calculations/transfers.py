"""Перемещение остатков с центрального склада в магазины до заказа поставщику.

Логика:
1. Считаем потребность магазина (recommended_order).
2. Берём доступный остаток «Склад основной» / «Склад основной1».
3. Распределяем по магазинам: сначала Флагман, далее по убыванию продаж SKU в точке.
4. Уменьшаем заказ поставщику на величину перемещения; остаток склада и магазинов обновляем.
"""

from __future__ import annotations

import logging
import math
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd

from config.settings import SETTINGS
from data.store_utils import (
    NETWORK_STORE_LABEL,
    UNKNOWN_STORE_LABEL,
    canon_store_name,
    is_central_warehouse,
)
from utils.helpers import safe_str

logger = logging.getLogger("zakupki_forecast.transfers")

PRIORITY_STORE = "Флагман"
CENTRAL_STORE_LABEL = "Склад основной1"


def _is_retail_dest(store: object) -> bool:
    name = safe_str(store).strip()
    if not name or name in {NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL}:
        return False
    return not is_central_warehouse(name)


def _rank_key(store: str, sales_qty: float, sales_amount: float) -> Tuple[int, float, float, str]:
    """Флагман всегда первый; остальные — по объёму продаж SKU в точке."""
    canon = canon_store_name(store) or store
    is_flagman = 0 if canon.casefold() == PRIORITY_STORE.casefold() else 1
    return (is_flagman, -float(sales_qty or 0), -float(sales_amount or 0), canon.casefold())


def apply_central_warehouse_transfers(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Распределяет остаток центрального склада по магазинам с дефицитом.

    Возвращает (обновлённый df, таблица перемещений).
    Если нет центрального склада или розничных строк — df без изменений, пустые перемещения.
    """
    out = df.copy()
    empty_transfers = pd.DataFrame(
        columns=[
            "from_store",
            "to_store",
            "priority_rank",
            "sku",
            "sku_key",
            "name",
            "barcode",
            "uom",
            "transfer_qty",
            "need_before",
            "store_stock_before",
            "central_stock_before",
            "central_stock_after",
            "order_after",
            "sales_qty",
            "supplier_name",
            "purchase_price",
        ]
    )

    if out.empty or "store" not in out.columns or "sku_key" not in out.columns:
        out["transfer_in"] = 0.0
        out["transfer_out"] = 0.0
        out["order_before_transfer"] = out.get("recommended_order", 0.0)
        return out, empty_transfers

    out["store"] = out["store"].map(lambda x: canon_store_name(x) if safe_str(x).strip() else safe_str(x))
    out["transfer_in"] = 0.0
    out["transfer_out"] = 0.0
    out["order_before_transfer"] = out["recommended_order"].fillna(0).astype(float)
    out["stock_before_transfer"] = out["stock"].fillna(0).astype(float)

    central_mask = out["store"].map(is_central_warehouse)
    retail_mask = out["store"].map(_is_retail_dest)
    if not bool(central_mask.any()) or not bool(retail_mask.any()):
        logger.info("Перемещения: нет центрального склада или розничных точек — пропуск")
        return out, empty_transfers

    # Доступный остаток на ЦС по SKU (сумма, если несколько строк)
    central_stock: Dict[str, float] = (
        out.loc[central_mask]
        .groupby("sku_key", dropna=False)["stock"]
        .sum()
        .astype(float)
        .to_dict()
    )
    central_stock = {k: max(0.0, float(v or 0)) for k, v in central_stock.items() if max(0.0, float(v or 0)) > 0}
    if not central_stock:
        logger.info("Перемещения: на центральном складе нет положительных остатков")
        return out, empty_transfers

    round_up = bool(SETTINGS.get("round_order_up", True))
    transfers: List[dict] = []

    # Индексы строк для обновления
    for sku_key, available in list(central_stock.items()):
        if available <= 0:
            continue
        demand = out.index[retail_mask & (out["sku_key"] == sku_key) & (out["order_before_transfer"] > 0)].tolist()
        if not demand:
            continue

        ranked = sorted(
            demand,
            key=lambda idx: _rank_key(
                safe_str(out.at[idx, "store"]),
                float(out.at[idx, "sales_qty"] if "sales_qty" in out.columns else 0) or 0.0,
                float(out.at[idx, "sales_amount"] if "sales_amount" in out.columns else 0) or 0.0,
            ),
        )

        remaining = float(available)
        central_before = float(available)
        for rank, idx in enumerate(ranked, start=1):
            if remaining <= 0:
                break
            need = float(out.at[idx, "recommended_order"] or 0)
            if need <= 0:
                continue
            qty = min(remaining, need)
            if round_up and qty > 0:
                qty = float(math.floor(qty + 1e-9))
            else:
                qty = max(0.0, round(qty, 3))
            if qty <= 0:
                continue

            remaining -= qty
            out.at[idx, "transfer_in"] = float(out.at[idx, "transfer_in"] or 0) + qty
            out.at[idx, "stock"] = float(out.at[idx, "stock"] or 0) + qty
            out.at[idx, "recommended_order"] = max(0.0, need - qty)

            transfers.append(
                {
                    "from_store": CENTRAL_STORE_LABEL,
                    "to_store": safe_str(out.at[idx, "store"]),
                    "priority_rank": rank,
                    "sku": safe_str(out.at[idx, "sku"]),
                    "sku_key": sku_key,
                    "name": safe_str(out.at[idx, "name"]),
                    "barcode": safe_str(out.at[idx, "barcode"] if "barcode" in out.columns else ""),
                    "uom": safe_str(out.at[idx, "uom"] if "uom" in out.columns else ""),
                    "transfer_qty": qty,
                    "need_before": need,
                    "store_stock_before": float(out.at[idx, "stock_before_transfer"] or 0),
                    "central_stock_before": central_before,
                    "central_stock_after": remaining,
                    "order_after": float(out.at[idx, "recommended_order"] or 0),
                    "sales_qty": float(out.at[idx, "sales_qty"] if "sales_qty" in out.columns else 0) or 0.0,
                    "supplier_name": safe_str(
                        out.at[idx, "supplier_name"] if "supplier_name" in out.columns else ""
                    ),
                    "purchase_price": float(
                        out.at[idx, "purchase_price"] if "purchase_price" in out.columns else 0
                    )
                    or 0.0,
                }
            )

        moved = central_before - remaining
        if moved <= 0:
            continue

        # Списываем с строк центрального склада (пропорционально остатку на строке)
        c_idx = out.index[central_mask & (out["sku_key"] == sku_key)].tolist()
        left_to_take = moved
        for idx in c_idx:
            if left_to_take <= 0:
                break
            row_stock = float(out.at[idx, "stock"] or 0)
            if row_stock <= 0:
                continue
            take = min(row_stock, left_to_take)
            out.at[idx, "stock"] = row_stock - take
            out.at[idx, "transfer_out"] = float(out.at[idx, "transfer_out"] or 0) + take
            left_to_take -= take

        # Пересчёт заказа на ЦС после списания (мин. остаток / потребность)
        min_stock_target = float(SETTINGS.get("min_stock_target", 0) or 0)
        for idx in c_idx:
            stock_now = float(out.at[idx, "stock"] or 0)
            need = float(out.at[idx, "need"] if "need" in out.columns else 0) or 0.0
            raw = max(0.0, need - stock_now)
            min_gap = max(0.0, min_stock_target - stock_now) if min_stock_target > 0 else 0.0
            blocked = bool(out.at[idx, "order_blocked"]) if "order_blocked" in out.columns else False
            if blocked:
                qty = min_gap
            else:
                qty = max(raw, min_gap, float(SETTINGS.get("min_order_qty", 0) or 0))
                max_mult = float(SETTINGS.get("max_order_multiplier", 5) or 5)
                forecast_adj = float(out.at[idx, "forecast_adj"] if "forecast_adj" in out.columns else 0) or 0.0
                if forecast_adj > 0:
                    qty = min(qty, forecast_adj * max_mult)
                qty = max(qty, min_gap)
            if round_up and qty > 0:
                qty = float(math.ceil(qty))
            else:
                qty = max(0.0, round(qty, 3))
            out.at[idx, "recommended_order"] = qty

    # Покрытие после перемещения
    if "avg_daily_sales" in out.columns:
        out["cover_days"] = np.where(
            out["avg_daily_sales"] > 0,
            out["stock"] / out["avg_daily_sales"],
            np.where(out["stock"] > 0, 9999.0, 0.0),
        )

    out["order_need_flag"] = out["recommended_order"] > 0
    transfers_df = pd.DataFrame(transfers) if transfers else empty_transfers
    if not transfers_df.empty:
        transfers_df = transfers_df.sort_values(
            ["sku_key", "priority_rank", "to_store"],
            ascending=[True, True, True],
        ).reset_index(drop=True)

    logger.info(
        "Перемещения со склада: строк=%s, qty=%s",
        len(transfers_df),
        float(transfers_df["transfer_qty"].sum()) if not transfers_df.empty else 0.0,
    )
    return out, transfers_df
