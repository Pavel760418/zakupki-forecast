"""Loading Excel stock/sales files via robust parser layer."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Tuple

import pandas as pd

from data.excel_parser import ParserError, parse_sales_file, parse_stock_file

logger = logging.getLogger("zakupki_forecast.loaders")


def load_stock_file(path: str | Path) -> pd.DataFrame:
    """
    Загружает файл остатков.
    Возвращает DataFrame: sku, name, stock, uom, store, barcode, sku_key.
    """
    path = Path(path)
    logger.info("Загрузка остатков: %s", path.name)
    try:
        out, diag = parse_stock_file(path)
    except ParserError as exc:
        raise ValueError(str(exc)) from exc

    logger.info("Остатки: %s позиций", len(out))
    logger.info("Диагностика парсинга: %s", diag.as_log_message())
    return out.reset_index(drop=True)


def load_sales_file(path: str | Path) -> pd.DataFrame:
    """
    Загружает файл продаж.
    Возвращает DataFrame: sku, name, date, qty, amount, store, barcode, sku_key.
    """
    path = Path(path)
    logger.info("Загрузка продаж: %s", path.name)
    try:
        out, diag = parse_sales_file(path)
    except ParserError as exc:
        raise ValueError(str(exc)) from exc

    if out.empty:
        raise ValueError("В файле продаж нет строк с корректной датой.")

    logger.info(
        "Продажи: %s строк, период %s — %s",
        len(out),
        out["date"].min().date(),
        out["date"].max().date(),
    )
    logger.info("Диагностика парсинга: %s", diag.as_log_message())
    return out.reset_index(drop=True)


def detect_sales_date_range(sales: pd.DataFrame) -> Tuple[pd.Timestamp, pd.Timestamp]:
    """Возвращает min/max дату продаж для подсказки периода."""
    return sales["date"].min(), sales["date"].max()
