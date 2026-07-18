"""Расчёт рекомендуемого заказа и статусов позиции."""

from __future__ import annotations

import math

import numpy as np
import pandas as pd

from config.settings import SETTINGS


def _abc_mult(abc: str) -> float:
    mapping = {
        "A": SETTINGS["abc_order_mult_a"],
        "B": SETTINGS["abc_order_mult_b"],
        "C": SETTINGS["abc_order_mult_c"],
    }
    return mapping.get(str(abc).upper(), 1.0)


def _safety_days(abc: str) -> float:
    mapping = {
        "A": SETTINGS["safety_stock_days_a"],
        "B": SETTINGS["safety_stock_days_b"],
        "C": SETTINGS["safety_stock_days_c"],
    }
    return float(mapping.get(str(abc).upper(), 1.0))


def apply_reorder_logic(
    df: pd.DataFrame,
    order_days: int,
    order_coef: float | None = None,
    uplift: float | None = None,
    downlift: float | None = None,
) -> pd.DataFrame:
    """
    Слои расчёта заказа:
    1) прогноз с трендом
    2) ABC-множитель
    3) коэффициенты заказа / uplift / снижение
    4) safety stock
    5) коррекция по остатку
    6) блок заказа для неликвидов
    """
    out = df.copy()
    order_coef = SETTINGS["order_coefficient"] if order_coef is None else float(order_coef)
    uplift = SETTINGS["uplift_coefficient"] if uplift is None else float(uplift)
    downlift = SETTINGS["downlift_coefficient"] if downlift is None else float(downlift)
    order_days = max(int(order_days), 1)

    out["abc_mult"] = out["abc_class"].map(_abc_mult)
    out["safety_days"] = out["abc_class"].map(_safety_days)
    out["order_coef"] = order_coef
    out["uplift_coef"] = uplift
    out["downlift_coef"] = downlift

    # Скорректированный прогноз
    out["forecast_adj"] = (
        out["forecast_base"]
        * out["abc_mult"]
        * out["order_coef"]
        * out["uplift_coef"]
        * out["downlift_coef"]
    )

    out["safety_stock"] = out["avg_daily_sales"] * out["safety_days"] * out["trend_coef"]
    out["need"] = out["forecast_adj"] + out["safety_stock"]
    out["raw_order"] = out["need"] - out["stock"]

    # Покрытие остатком в днях
    out["cover_days"] = np.where(
        out["avg_daily_sales"] > 0,
        out["stock"] / out["avg_daily_sales"],
        np.where(out["stock"] > 0, 9999.0, 0.0),
    )

    # Блокировка неликвидов
    dead_days = SETTINGS["dead_stock_days_no_sales"]
    is_dead = (out["sales_qty"] <= 0) | (out["days_since_sale"] >= dead_days)
    out["is_dead_stock"] = is_dead & (out["stock"] > 0)
    out["is_no_movement"] = out["sales_qty"] <= 0

    out["order_blocked"] = False
    if SETTINGS["block_order_dead_stock"]:
        # Блокируем заказ, если нет продаж и нет растущего тренда
        block_mask = out["is_no_movement"] & (out["trend_coef"] <= 1.0)
        out.loc[block_mask, "order_blocked"] = True

    # Рекомендуемый заказ
    max_mult = SETTINGS["max_order_multiplier"]
    orders = []
    for _, row in out.iterrows():
        if row["order_blocked"]:
            orders.append(0.0)
            continue
        qty = max(float(row["raw_order"]), float(SETTINGS["min_order_qty"]))
        cap = float(row["forecast_adj"]) * max_mult if row["forecast_adj"] > 0 else qty
        if cap > 0:
            qty = min(qty, cap)
        if SETTINGS["round_order_up"] and qty > 0:
            qty = float(math.ceil(qty))
        else:
            qty = max(0.0, round(qty, 3))
        orders.append(qty)

    out["recommended_order"] = orders
    out["order_need_flag"] = out["recommended_order"] > 0
    return out
