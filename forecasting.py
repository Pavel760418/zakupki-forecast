"""Прогноз спроса: средние продажи + тренд."""

from __future__ import annotations

import numpy as np
import pandas as pd

from config.settings import SETTINGS


def calc_trend_coefficient(sales_h1: float, sales_h2: float) -> float:
    """
    Тренд = продажи 2-й половины / 1-й половины периода.
    Ограничивается [trend_min, trend_max].
    """
    h1 = float(sales_h1 or 0)
    h2 = float(sales_h2 or 0)
    if h1 <= 0 and h2 <= 0:
        return SETTINGS["trend_neutral"]
    if h1 <= 0 and h2 > 0:
        return SETTINGS["trend_max"]
    trend = h2 / h1
    return float(np.clip(trend, SETTINGS["trend_min"], SETTINGS["trend_max"]))


def apply_forecast_metrics(df: pd.DataFrame, period_days: int, order_days: int) -> pd.DataFrame:
    """
    Базовые метрики прогноза (числовые seed-значения для Excel).
    Формулы в Excel будут пересчитывать то же самое от редактируемых ячеек.
    """
    out = df.copy()
    period_days = max(int(period_days), 1)
    order_days = max(int(order_days), 1)

    out["avg_daily_sales"] = out["sales_qty"] / period_days
    out["trend_coef"] = [
        calc_trend_coefficient(h1, h2)
        for h1, h2 in zip(out["sales_h1"], out["sales_h2"])
    ]

    # Базовый прогноз на горизонт заказа с учётом тренда
    out["forecast_base"] = out["avg_daily_sales"] * order_days * out["trend_coef"]

    # Направление тренда для аналитики
    def _trend_label(t: float) -> str:
        if t >= 1.15:
            return "Рост"
        if t <= 0.85:
            return "Спад"
        return "Стабильно"

    out["trend_label"] = out["trend_coef"].map(_trend_label)
    return out
