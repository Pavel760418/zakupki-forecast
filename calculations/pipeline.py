"""Оркестрация всех слоёв расчёта."""

from __future__ import annotations

import logging
from typing import Any, Dict, Tuple

import pandas as pd

from calculations.abc_analysis import apply_abc_analysis
from calculations.forecasting import apply_forecast_metrics
from calculations.quantum_orders import (
    apply_central_supplier_order_from_stores,
    attach_quantum_column,
    round_orders_to_quantum,
)
from calculations.reorder import apply_reorder_logic
from calculations.risk_analysis import apply_risk_analysis
from calculations.transfers import apply_central_warehouse_transfers
from config.settings import SETTINGS
from data.merge import GRAIN_NETWORK, GRAIN_STORE, build_product_frame, filter_sales_by_period
from data.store_utils import has_retail_store_dimension, has_store_dimension
from data.store_utils import NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL
from data.supplier_mapping import attach_supplier_attributes
from data.validators import validate_frames, validate_period

logger = logging.getLogger("zakupki_forecast.pipeline")


def _apply_abc_network_then_map(df: pd.DataFrame) -> pd.DataFrame:
    """ABC всегда по SKU в целом по сети, затем класс копируется на строки магазинов."""
    sku_base = (
        df.groupby("sku_key", as_index=False)
        .agg(
            sales_qty=("sales_qty", "sum"),
            sales_amount=("sales_amount", "sum"),
            name=("name", "first"),
            sku=("sku", "first"),
        )
    )
    abc = apply_abc_analysis(sku_base)
    cols = [c for c in ("sku_key", "abc_class", "abc_share", "abc_cum_share", "abc_rank") if c in abc.columns]
    out = df.drop(columns=[c for c in ("abc_class", "abc_share", "abc_cum_share", "abc_rank") if c in df.columns])
    return out.merge(abc[cols], on="sku_key", how="left")


def run_calculations(
    stock: pd.DataFrame,
    sales: pd.DataFrame,
    date_from,
    date_to,
    order_period_days: int | None = None,
    order_coefficient: float | None = None,
    uplift_coefficient: float | None = None,
    downlift_coefficient: float | None = None,
    allow_empty_sales: bool = False,
    grain: str = GRAIN_NETWORK,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Полный пайплайн:
    валидация → фильтр продаж → витрина → ABC → прогноз → заказ → риски.

    allow_empty_sales=True — режим поставщика без продаж по его SKU:
    расчёт идёт по ассортименту/остаткам (в т.ч. нулевым) без падения.
    grain: network (сводно) или store (по магазинам).
    """
    if allow_empty_sales and (sales is None or sales.empty):
        sales = pd.DataFrame(columns=["sku", "name", "date", "qty", "amount", "sku_key"])
        if stock is None or stock.empty:
            raise ValueError("Таблица остатков пуста.")
        for col in ("sku", "name", "stock", "sku_key"):
            if col not in stock.columns:
                raise ValueError(f"В остатках нет колонки: {col}")
        d1, d2, period_days = validate_period(date_from, date_to)
    else:
        validate_frames(stock, sales)
        d1, d2, period_days = validate_period(
            date_from,
            date_to,
            sales_min=sales["date"].min(),
            sales_max=sales["date"].max(),
        )
    order_days = int(order_period_days or SETTINGS["default_order_period_days"])

    requested_grain = GRAIN_STORE if grain == GRAIN_STORE else GRAIN_NETWORK
    store_fallback = False
    stock_no_retail = False
    sales_has_stores = has_store_dimension(sales)
    stock_has_stores = has_store_dimension(stock)
    stock_has_retail = has_retail_store_dimension(stock)

    if requested_grain == GRAIN_STORE:
        # Достаточно магазинов в продажах ИЛИ в остатках.
        # Типичный кейс 1С: продажи по подразделениям, остатки только «Склад основной1».
        if not sales_has_stores and not stock_has_stores:
            logger.warning("Нет магазинов ни в остатках, ни в продажах — откат на сводный расчёт")
            requested_grain = GRAIN_NETWORK
            store_fallback = True
        elif sales_has_stores and not stock_has_retail:
            stock_no_retail = True
            logger.info(
                "Магазины взяты из продаж; в остатках нет розничных точек "
                "(часто только центральный склад) — grain=store без разноски остатков"
            )

    sales_period = filter_sales_by_period(sales, d1, d2)
    base, meta = build_product_frame(stock, sales_period, d1, d2, grain=requested_grain)

    if requested_grain == GRAIN_STORE:
        df = _apply_abc_network_then_map(base)
    else:
        df = apply_abc_analysis(base)

    df = apply_forecast_metrics(df, period_days, order_days)
    df = apply_reorder_logic(
        df,
        order_days=order_days,
        order_coef=order_coefficient,
        uplift=uplift_coefficient,
        downlift=downlift_coefficient,
    )
    df = attach_quantum_column(df)
    df = round_orders_to_quantum(df, "recommended_order")

    # Поставщик/цена → перемещения квантами с ЦС → заказ на ЦС = сумма магазинов → риски.
    df = attach_supplier_attributes(df)
    transfers_df = None
    if requested_grain == GRAIN_STORE:
        df, transfers_df = apply_central_warehouse_transfers(df)
        df = round_orders_to_quantum(df, "recommended_order")
        df = apply_central_supplier_order_from_stores(df)
    else:
        df["transfer_in"] = 0.0
        df["transfer_out"] = 0.0
        df["order_before_transfer"] = df["recommended_order"].fillna(0)
        df["supplier_order_qty"] = df["recommended_order"].fillna(0)

    df = apply_risk_analysis(df)
    # Служебные строки без магазина не заказываем (ошибка сборки ассортимента).
    if "store" in df.columns:
        bad = df["store"].fillna("").isin({"", UNKNOWN_STORE_LABEL, NETWORK_STORE_LABEL})
        if bad.any():
            df.loc[bad, "recommended_order"] = 0.0
            if "supplier_order_qty" in df.columns:
                df.loc[bad, "supplier_order_qty"] = 0.0
            df.loc[bad, "order_need_flag"] = False

    # Сумма заказа поставщику: в режиме магазинов — только ЦС (supplier_order_qty).
    if "supplier_order_qty" in df.columns:
        df["order_sum"] = df["supplier_order_qty"].fillna(0) * df["purchase_price"].fillna(0)
    else:
        df["order_sum"] = df["recommended_order"].fillna(0) * df["purchase_price"].fillna(0)

    if "store" in df.columns:
        df = df.sort_values(
            ["priority", "store", "abc_class", "recommended_order"],
            ascending=[True, True, True, False],
        )
    else:
        df = df.sort_values(
            ["priority", "abc_class", "recommended_order"],
            ascending=[True, True, False],
        )
    df = df.reset_index(drop=True)
    df["row_id"] = range(1, len(df) + 1)

    transfer_lines = int(len(transfers_df)) if transfers_df is not None and not transfers_df.empty else 0
    transfer_qty = (
        float(transfers_df["transfer_qty"].sum())
        if transfers_df is not None and not transfers_df.empty
        else 0.0
    )

    if requested_grain == GRAIN_STORE and "supplier_order_qty" in df.columns:
        order_mask = df["supplier_order_qty"].fillna(0) > 0
        order_lines = int(order_mask.sum())
        order_qty_total = float(df.loc[order_mask, "supplier_order_qty"].sum())
        order_sum_total = float(df.loc[order_mask, "order_sum"].sum())
    else:
        order_lines = int((df["recommended_order"] > 0).sum())
        order_qty_total = float(df["recommended_order"].sum())
        order_sum_total = float(df["order_sum"].sum())

    meta.update(
        {
            "order_period_days": order_days,
            "order_coefficient": (
                float(order_coefficient)
                if order_coefficient is not None
                else float(SETTINGS["order_coefficient"])
            ),
            "uplift_coefficient": (
                float(uplift_coefficient)
                if uplift_coefficient is not None
                else float(SETTINGS["uplift_coefficient"])
            ),
            "downlift_coefficient": (
                float(downlift_coefficient)
                if downlift_coefficient is not None
                else float(SETTINGS["downlift_coefficient"])
            ),
            "min_stock_target": SETTINGS.get("min_stock_target", 24),
            "abc_a_count": int((df["abc_class"] == "A").sum()),
            "abc_b_count": int((df["abc_class"] == "B").sum()),
            "abc_c_count": int((df["abc_class"] == "C").sum()),
            "oos_count": int(df["is_oos_risk"].sum()),
            "overstock_count": int(df["is_overstock"].sum()),
            "dead_count": int(df["is_dead_stock"].sum()),
            "order_lines": order_lines,
            "order_qty_total": order_qty_total,
            "order_sum_total": order_sum_total,
            "grain": requested_grain,
            "store_grain_fallback": store_fallback,
            "stock_no_retail_stores": stock_no_retail,
            "sales_has_stores": sales_has_stores,
            "stock_has_stores": stock_has_stores,
            "transfer_lines": transfer_lines,
            "transfer_qty_total": transfer_qty,
            "transfers": transfers_df,
            "quantum_enabled": True,
            "release": "4",
        }
    )
    logger.info(
        "Расчёт завершён: строк=%s, к заказу=%s, OOS=%s, grain=%s",
        meta["items_count"],
        meta["order_lines"],
        meta["oos_count"],
        requested_grain,
    )
    return df, meta
