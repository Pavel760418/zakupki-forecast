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
    p = argparse.ArgumentParser(description="Автозаказ СИКС (CLI)")
    p.add_argument("--stock", required=True, help="Путь к Excel остатков")
    p.add_argument("--sales", required=True, help="Путь к Excel продаж")
    p.add_argument("--from", dest="date_from", required=True, help="Дата начала ДД.ММ.ГГГГ")
    p.add_argument("--to", dest="date_to", required=True, help="Дата окончания ДД.ММ.ГГГГ")
    p.add_argument("--order-days", type=int, default=14, help="Горизонт заказа, дни")
    p.add_argument("--order-coef", type=float, default=1.0, help="Коэффициент заказа")
    p.add_argument(
        "--supplier",
        default=None,
        help="Опционально: имя поставщика из файла привязки (иначе общий расчёт)",
    )
    p.add_argument("--out", default=None, help="Путь итогового xlsx")
    return p.parse_args()


def main() -> int:
    setup_logging()
    args = parse_args()
    ensure_output_dir()

    stock = load_stock_file(args.stock)
    sales = load_sales_file(args.sales)

    supplier_mode = False
    supplier_name = ""
    if args.supplier:
        from data.supplier_mapping import filter_frames_by_supplier, load_supplier_mapping

        mapping_result = load_supplier_mapping()
        if not mapping_result.loaded:
            print("WARN: файл привязки недоступен — выполняется общий расчёт", file=sys.stderr)
        else:
            stock, sales, info = filter_frames_by_supplier(
                stock, sales, mapping_result.mapping, args.supplier
            )
            supplier_mode = True
            supplier_name = args.supplier
            print(f"Фильтр поставщика: {info}")

    df, meta = run_calculations(
        stock,
        sales,
        args.date_from,
        args.date_to,
        order_period_days=args.order_days,
        order_coefficient=args.order_coef,
    )
    meta["supplier_mode"] = supplier_mode
    meta["supplier_name"] = supplier_name if supplier_mode else ""
    path = build_workbook(df, meta, args.out)
    print(f"OK: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
