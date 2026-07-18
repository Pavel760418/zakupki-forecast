"""Оркестрация всех слоёв расчёта."""

from __future__ import annotations

import logging
from typing import Any, Dict, Tuple

import pandas as pd

from calculations.abc_analysis import apply_abc_analysis
from calculations.forecasting import apply_forecast_metrics
from calculations.reorder import apply_reorder_logic
from calculations.risk_analysis import apply_risk_analysis
from config.settings import SETTINGS
from data.merge import build_product_frame, filter_sales_by_period
from data.validators import validate_frames, validate_period

logger = logging.getLogger("zakupki_forecast.pipeline")


def run_calculations(
    stock: pd.DataFrame,
    sales: pd.DataFrame,
    date_from,
    date_to,
    order_period_days: int | None = None,
    order_coefficient: float | None = None,
    uplift_coefficient: float | None = None,
    downlift_coefficient: float | None = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Полный пайплайн:
    валидация → фильтр продаж → витрина → ABC → прогноз → заказ → риски.
    """
    validate_frames(stock, sales)
    d1, d2, period_days = validate_period(
        date_from,
        date_to,
        sales_min=sales["date"].min(),
        sales_max=sales["date"].max(),
    )
    order_days = int(order_period_days or SETTINGS["default_order_period_days"])

    sales_period = filter_sales_by_period(sales, d1, d2)
    base, meta = build_product_frame(stock, sales_period, d1, d2)

    df = apply_abc_analysis(base)
    df = apply_forecast_metrics(df, period_days, order_days)
    df = apply_reorder_logic(
        df,
        order_days=order_days,
        order_coef=order_coefficient,
        uplift=uplift_coefficient,
        downlift=downlift_coefficient,
    )
    df = apply_risk_analysis(df)

    df = df.sort_values(["priority", "abc_class", "recommended_order"], ascending=[True, True, False])
    df = df.reset_index(drop=True)
    df["row_id"] = range(1, len(df) + 1)

    meta.update(
        {
            "order_period_days": order_days,
            "order_coefficient": order_coefficient or SETTINGS["order_coefficient"],
            "uplift_coefficient": uplift_coefficient or SETTINGS["uplift_coefficient"],
            "downlift_coefficient": downlift_coefficient or SETTINGS["downlift_coefficient"],
            "abc_a_count": int((df["abc_class"] == "A").sum()),
            "abc_b_count": int((df["abc_class"] == "B").sum()),
            "abc_c_count": int((df["abc_class"] == "C").sum()),
            "oos_count": int(df["is_oos_risk"].sum()),
            "overstock_count": int(df["is_overstock"].sum()),
            "dead_count": int(df["is_dead_stock"].sum()),
            "order_lines": int((df["recommended_order"] > 0).sum()),
            "order_qty_total": float(df["recommended_order"].sum()),
        }
    )
    logger.info(
        "Расчёт завершён: SKU=%s, к заказу=%s, OOS=%s",
        meta["items_count"],
        meta["order_lines"],
        meta["oos_count"],
    )
    return df, meta
