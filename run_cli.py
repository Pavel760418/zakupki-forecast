"""
CLI-режим без GUI (для отладки в PyCharm / скриптов).

Пример:
    python run_cli.py --stock samples/остатки.xlsx --sales samples/продажи.xlsx ^
        --from 01.06.2026 --to 30.06.2026 --order-days 14
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from calculations.pipeline import run_calculations
from data.loaders import load_sales_file, load_stock_file
from excel.workbook_builder import build_workbook
from utils.helpers import ensure_output_dir
from utils.logging_config import setup_logging


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Прогноз заказа (CLI)")
    p.add_argument("--stock", required=True, help="Путь к Excel остатков")
    p.add_argument("--sales", required=True, help="Путь к Excel продаж")
    p.add_argument("--from", dest="date_from", required=True, help="Дата начала ДД.ММ.ГГГГ")
    p.add_argument("--to", dest="date_to", required=True, help="Дата окончания ДД.ММ.ГГГГ")
    p.add_argument("--order-days", type=int, default=14, help="Горизонт заказа, дни")
    p.add_argument("--order-coef", type=float, default=1.0, help="Коэффициент заказа")
    p.add_argument("--out", default=None, help="Путь итогового xlsx")
    return p.parse_args()


def main() -> int:
    setup_logging()
    args = parse_args()
    ensure_output_dir()

    stock = load_stock_file(args.stock)
    sales = load_sales_file(args.sales)
    df, meta = run_calculations(
        stock,
        sales,
        args.date_from,
        args.date_to,
        order_period_days=args.order_days,
        order_coefficient=args.order_coef,
    )
    path = build_workbook(df, meta, args.out)
    print(f"OK: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
