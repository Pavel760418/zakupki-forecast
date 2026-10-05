"""Прогноз новых точек формата Флагман: Лига Химки и Сколково.

Если своей истории продаж нет или она меньше доли продаж Флагмана по тому же SKU,
среднедневные и прогноз берутся с Флагмана.
Рекомендуемый заказ = потребность Флагмана − собственный остаток точки.
Чужой склад в эту потребность не входит. Мин. остаток сети поверх аналога не докупается:
если остаток уже покрывает потребность Флагмана, заказ 0.
"""

from __future__ import annotations

import logging

import numpy as np
import pandas as pd

from config.settings import SETTINGS
from data.store_utils import BENCHMARK_STORE, is_flagman_analog_store

logger = logging.getLogger("zakupki_forecast.flagman_analog")


def _num(value: object, default: float = 0.0) -> float:
    try:
        number = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return default
    if number != number:
        return default
    return number


def _own_history_is_thin(own_sales: float, benchmark_sales: float, share_max: float) -> bool:
    own = float(own_sales or 0)
    bench = float(benchmark_sales or 0)
    if own <= 0:
        return True
    if bench <= 0:
        return False
    return own < float(share_max) * bench


def apply_flagman_analog(df: pd.DataFrame) -> pd.DataFrame:
    """Подменяет прогноз и заказ Лиги Химки / Сколково аналогом Флагмана по SKU."""
    out = df.copy()
    if out.empty or "store" not in out.columns or "sku_key" not in out.columns:
        return out

    share_max = float(SETTINGS.get("analog_sales_share_max", 1.0) or 1.0)
    out["analog_of_flagman"] = False
    out["analog_note"] = ""

    bench_mask = out["store"].map(lambda x: (x or "") == BENCHMARK_STORE)
    analog_mask = out["store"].map(is_flagman_analog_store)
    if not bool(bench_mask.any()) or not bool(analog_mask.any()):
        return out

    bench = (
        out.loc[bench_mask]
        .groupby("sku_key", as_index=False)
        .agg(
            bench_sales=("sales_qty", "sum"),
            bench_avg=("avg_daily_sales", "sum"),
            bench_trend=("trend_coef", "max"),
            bench_forecast_base=("forecast_base", "sum"),
            bench_forecast_adj=("forecast_adj", "sum"),
            bench_safety=("safety_stock", "sum"),
            bench_need=("need", "sum"),
        )
    )
    bench_by_sku = bench.set_index("sku_key")

    applied = 0
    for idx in out.index[analog_mask]:
        sku_key = out.at[idx, "sku_key"]
        if sku_key not in bench_by_sku.index:
            out.at[idx, "analog_note"] = "Нет аналога на Флагмане — оставлен расчёт точки"
            continue
        row = bench_by_sku.loc[sku_key]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        own_sales = _num(out.at[idx, "sales_qty"])
        bench_sales = _num(row["bench_sales"])
        if not _own_history_is_thin(own_sales, bench_sales, share_max):
            out.at[idx, "analog_note"] = (
                "Свои продажи не меньше доли Флагмана — оставлен расчёт точки"
            )
            continue

        stock = _num(out.at[idx, "stock"])
        need = _num(row["bench_need"])
        order = max(0.0, need - stock)
        avg = _num(row["bench_avg"])

        out.at[idx, "avg_daily_sales"] = avg
        out.at[idx, "trend_coef"] = _num(row["bench_trend"], 1.0)
        out.at[idx, "forecast_base"] = _num(row["bench_forecast_base"])
        out.at[idx, "forecast_adj"] = _num(row["bench_forecast_adj"])
        out.at[idx, "safety_stock"] = _num(row["bench_safety"])
        out.at[idx, "need"] = need
        out.at[idx, "raw_order"] = order
        out.at[idx, "recommended_order"] = order
        out.at[idx, "order_need_flag"] = order > 0
        out.at[idx, "order_blocked"] = False
        if bench_sales > 0:
            out.at[idx, "is_dead_stock"] = False
        if avg > 0:
            out.at[idx, "cover_days"] = stock / avg
        else:
            out.at[idx, "cover_days"] = 9999.0 if stock > 0 else 0.0
        out.at[idx, "analog_of_flagman"] = True
        if order <= 0:
            out.at[idx, "analog_note"] = (
                f"Аналог Флагмана {need:.0f} шт уже покрыт остатком {stock:.0f}"
            )
        else:
            out.at[idx, "analog_note"] = (
                f"Заказ как у Флагмана: потребность {need:.0f} − остаток {stock:.0f}"
            )
        applied += 1

    if "cover_days" in out.columns:
        out["cover_days"] = pd.to_numeric(out["cover_days"], errors="coerce").fillna(0.0)
        out["cover_days"] = np.where(out["cover_days"] < 0, 0.0, out["cover_days"])

    logger.info("Аналог Флагмана: строк с подменой прогноза/заказа=%s", applied)
    return out
