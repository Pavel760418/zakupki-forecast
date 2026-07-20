"""
Публичный API модуля.
Основная точка входа для интеграции: ZAPriceCalculator.run(...)
"""
from __future__ import annotations

import logging
from pathlib import Path

from za_price_calculator.builder import build_workbook, save_workbook
from za_price_calculator.io_handlers.loader import load_sales_file, load_source_file

logger = logging.getLogger(__name__)


class ZAPriceCalculator:
    """
    Основной сервис-класс калькулятора цен «Зеленое Яблоко».

    Использование:
        calc = ZAPriceCalculator()
        result_path = calc.run("прайс.xlsx", output_path="output/ZA_Price_Calculator.xlsx")

    С опциональным файлом продаж:
        result_path = calc.run("прайс.xlsx", sales_path="продажи.xlsx", output_path="out.xlsx")
    """

    def run(
        self,
        source_path,
        output_path,
        sales_path=None,
        source_sheet=0,
        sales_sheet=0,
    ) -> Path:
        """
        Полный цикл: загрузка -> валидация -> построение книги -> сохранение.

        :param source_path: путь к обязательному Excel-файлу прайс-листа.
        :param output_path: путь для сохранения итогового файла.
        :param sales_path: опциональный путь к Excel-файлу продаж.
        :param source_sheet: имя/индекс листа прайс-листа во входном файле.
        :param sales_sheet: имя/индекс листа продаж во входном файле.
        :return: Path сохранённого итогового файла.
        """
        logger.info("Запуск расчёта: source=%s, sales=%s", source_path, sales_path)

        source_df = load_source_file(source_path, sheet_name=source_sheet)

        sales_df = None
        if sales_path is not None:
            sales_df = load_sales_file(sales_path, sheet_name=sales_sheet)

        workbook = build_workbook(source_df, sales_df)
        saved_path = save_workbook(workbook, output_path)

        logger.info("Расчёт завершён успешно: %s", saved_path)
        return saved_path
