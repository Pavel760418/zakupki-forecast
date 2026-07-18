"""Валидация входных данных и параметров периода."""

from __future__ import annotations

from datetime import date, datetime
from typing import Tuple, Union

import pandas as pd

from config.settings import SETTINGS


DateLike = Union[str, date, datetime, pd.Timestamp]


def parse_date(value: DateLike) -> pd.Timestamp:
    """Парсит дату пользователя."""
    ts = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        raise ValueError(f"Некорректная дата: {value}")
    return pd.Timestamp(ts).normalize()


def validate_period(
    date_from: DateLike,
    date_to: DateLike,
    sales_min: pd.Timestamp | None = None,
    sales_max: pd.Timestamp | None = None,
) -> Tuple[pd.Timestamp, pd.Timestamp, int]:
    """
    Проверяет период расчёта.
    Возвращает (date_from, date_to, days_inclusive).
    """
    d1 = parse_date(date_from)
    d2 = parse_date(date_to)
    if d2 < d1:
        raise ValueError("Дата окончания периода раньше даты начала.")

    days = int((d2 - d1).days) + 1
    if days < SETTINGS["min_period_days"] or days > SETTINGS["max_period_days"]:
        raise ValueError(
            f"Период должен быть от {SETTINGS['min_period_days']} "
            f"до {SETTINGS['max_period_days']} дней (сейчас {days})."
        )

    if sales_min is not None and sales_max is not None:
        # Период должен пересекаться с данными продаж
        if d2 < sales_min or d1 > sales_max:
            raise ValueError(
                "Указанный период не пересекается с датами в файле продаж "
                f"({sales_min.date()} — {sales_max.date()})."
            )

    return d1, d2, days


def validate_frames(stock: pd.DataFrame, sales: pd.DataFrame) -> None:
    """Базовые проверки датафреймов перед расчётом."""
    if stock is None or stock.empty:
        raise ValueError("Таблица остатков пуста.")
    if sales is None or sales.empty:
        raise ValueError("Таблица продаж пуста.")
    for col in ("sku", "name", "stock", "sku_key"):
        if col not in stock.columns:
            raise ValueError(f"В остатках нет колонки: {col}")
    for col in ("sku", "name", "date", "qty", "sku_key"):
        if col not in sales.columns:
            raise ValueError(f"В продажах нет колонки: {col}")
