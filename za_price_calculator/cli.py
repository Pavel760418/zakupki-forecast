"""
CLI-точка входа для запуска в PyCharm / из терминала.

Использование:
    python -m za_price_calculator.cli --source прайс.xlsx --output out.xlsx
    python -m za_price_calculator.cli --source прайс.xlsx --sales продажи.xlsx --output out.xlsx
"""
from __future__ import annotations

import argparse
import logging
import sys

from za_price_calculator.exceptions import ZAPriceCalculatorError
from za_price_calculator.service import ZAPriceCalculator


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="za_price_calculator",
        description="Автоматический калькулятор цен, наценки, маржи и сценариев Зеленое Яблоко",
    )
    parser.add_argument("--source", required=True, help="Путь к Excel-файлу с прайс-листом (обязательно)")
    parser.add_argument("--output", required=True, help="Путь для сохранения итогового файла")
    parser.add_argument("--sales", required=False, default=None, help="Путь к Excel-файлу с продажами (опционально)")
    parser.add_argument("--source-sheet", required=False, default=0, help="Имя/индекс листа прайс-листа")
    parser.add_argument("--sales-sheet", required=False, default=0, help="Имя/индекс листа продаж")
    parser.add_argument("-v", "--verbose", action="store_true", help="Подробный вывод логов")
    return parser


def main(argv=None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    try:
        source_sheet = int(args.source_sheet) if str(args.source_sheet).isdigit() else args.source_sheet
        sales_sheet = int(args.sales_sheet) if str(args.sales_sheet).isdigit() else args.sales_sheet

        calc = ZAPriceCalculator()
        result_path = calc.run(
            source_path=args.source,
            output_path=args.output,
            sales_path=args.sales,
            source_sheet=source_sheet,
            sales_sheet=sales_sheet,
        )
        print(f"Готово. Файл сохранён: {result_path}")
        return 0
    except ZAPriceCalculatorError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Непредвиденная ошибка: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
