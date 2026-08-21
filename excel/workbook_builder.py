"""
Построение итогового Excel-файла с формулами.

Листы:
  00_Инструкция
  01_Настройки
  02_Дашборд
  03_Расчёт_заказа   ← основная рабочая таблица с формулами
  04_ABC
  05_Риск_OOS
  06_Неликвиды
  07_Избыток
  08_Тренды
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from config.settings import OUTPUT_DIR, SETTINGS
from excel.formatting import (
    ALIGN_CENTER,
    ALIGN_WRAP,
    FILL_EDIT,
    FILL_FORMULA,
    FILL_HEADER,
    FILL_READONLY,
    FONT_HEADER,
    THIN,
    add_abc_validation,
    add_oos_conditional,
    add_order_highlight,
    autosize_columns,
    style_header_row,
)
from utils.helpers import ensure_output_dir, safe_str, timestamp_filename

logger = logging.getLogger("zakupki_forecast.excel")

# Индексы колонок листа «Расчёт_заказа» (1-based)
# Редактируемые (жёлтые): C остаток, D продажи, E тренд ручной, F коэф.строки, G ABC
# Формульные (зелёные): остальное
COL = {
    "id": 1,
    "sku": 2,
    "name": 3,
    "stock": 4,
    "sales": 5,
    "trend": 6,
    "line_coef": 7,
    "abc": 8,
    "avg_daily": 9,
    "forecast": 10,
    "safety": 11,
    "need": 12,
    "raw_order": 13,
    "rec_order": 14,
    "cover": 15,
    "status": 16,
    "risk": 17,
    "light": 18,
    "uom": 19,
    "note": 20,
}

HEADERS = [
    "№",
    "Артикул",
    "Наименование",
    "Остаток",
    "Продажи за период",
    "Тренд (коэф.)",
    "Коэф. строки",
    "ABC",
    "Среднедневные",
    "Прогноз продаж",
    "Страховой запас",
    "Потребность",
    "Потребность − остаток",
    "Рекомендуемый заказ",
    "Покрытие, дни",
    "Статус",
    "Риск",
    "Светофор",
    "Ед.изм.",
    "Комментарий закупщика",
]


def build_workbook(
    df: pd.DataFrame,
    meta: Dict[str, Any],
    output_path: str | Path | None = None,
    supplier_name: str | None = None,
) -> Path:
    """Создаёт итоговый xlsx и возвращает путь.

    supplier_name — опционально. Если передан, добавляется лист «09_Заказ_поставщику».
    Без выбора поставщика структура книги остаётся прежней.
    """
    ensure_output_dir()
    if output_path is None:
        output_path = OUTPUT_DIR / timestamp_filename(SETTINGS["workbook_name_prefix"])
    else:
        output_path = Path(output_path)

    wb = Workbook()
    # Удаляем дефолтный лист — создадим свои
    default = wb.active
    wb.remove(default)

    _build_instruction(wb, supplier_name=supplier_name)
    _build_settings(wb, meta)
    _build_dashboard(wb, df, meta)
    _build_main_calc(wb, df, meta)
    _build_abc_sheet(wb, df)
    _build_filtered_sheet(wb, df[df["is_oos_risk"]], "05_Риск_OOS", "Риски out-of-stock")
    _build_filtered_sheet(wb, df[df["is_dead_stock"]], "06_Неликвиды", "Неликвиды и зависшие позиции")
    _build_filtered_sheet(wb, df[df["is_overstock"]], "07_Избыток", "Избыточные / завышенные запасы")
    _build_trends_sheet(wb, df)

    # Опциональный лист: только при явном выборе поставщика
    if supplier_name and safe_str(supplier_name).strip():
        order_view = df[df["recommended_order"] > 0].copy() if "recommended_order" in df.columns else df
        _build_supplier_order_sheet(wb, order_view, safe_str(supplier_name).strip())

    wb.save(output_path)
    logger.info("Excel сохранён: %s", output_path)
    return output_path


def _ws(wb: Workbook, title: str):
    return wb.create_sheet(title)


def _build_instruction(wb: Workbook, supplier_name: str | None = None) -> None:
    ws = _ws(wb, "00_Инструкция")
    lines = [
        ("ИНСТРУКЦИЯ ДЛЯ ЗАКУПЩИКА", True, 16),
        ("", False, 11),
        ("1. Назначение файла", True, 13),
        (
            "Файл помогает рассчитать прогноз продаж и рекомендуемый заказ на основе остатков и продаж из 1С. "
            "Главный рабочий лист — «03_Расчёт_заказа». Параметры меняйте на «01_Настройки».",
            False,
            11,
        ),
        ("", False, 11),
        ("2. Какие листы для чего", True, 13),
        ("00_Инструкция — эта памятка.", False, 11),
        ("01_Настройки — периоды, коэффициенты, пороги ABC/OOS/неликвидов. Меняйте жёлтые ячейки.", False, 11),
        ("02_Дашборд — сводные показатели и приоритеты.", False, 11),
        ("03_Расчёт_заказа — основная таблица: формулы пересчитывают заказ при изменении параметров.", False, 11),
        ("04_ABC — классификация товаров A/B/C.", False, 11),
        ("05_Риск_OOS — позиции с риском дефицита.", False, 11),
        ("06_Неликвиды — нет продаж / зависший остаток.", False, 11),
        ("07_Избыток — завышенные запасы.", False, 11),
        ("08_Тренды — рост / спад спроса.", False, 11),
        (
            "09_Заказ_поставщику — заказ только по выбранному поставщику (появляется при режиме расчёта по поставщику).",
            False,
            11,
        ),
    ]
    if supplier_name and safe_str(supplier_name).strip():
        lines.append(
            (
                f"В этой книге выбран поставщик: {safe_str(supplier_name).strip()}.",
                False,
                11,
            )
        )
    lines.extend(
        [
        ("", False, 11),
        ("3. Что МОЖНО редактировать (жёлтые ячейки)", True, 13),
        ("На «01_Настройки»: дни периода продаж, дни периода заказа, коэф. заказа, uplift/downlift, safety stock, пороги.", False, 11),
        ("На «03_Расчёт_заказа»: Остаток, Продажи за период, Тренд, Коэф. строки, ABC, Комментарий, Наименование.", False, 11),
        ("Можно добавлять строки ВНИЗУ таблицы, копируя формулы с соседней строки.", False, 11),
        ("", False, 11),
        ("4. Что НЕЛЬЗЯ редактировать (зелёные / расчётные)", True, 13),
        ("Среднедневные, Прогноз, Страховой запас, Потребность, Рекомендуемый заказ, Покрытие, Статус — считаются формулами.", False, 11),
        ("Не удаляйте лист «01_Настройки» и именованную логику ссылок вида Настройки!$B$…", False, 11),
        ("", False, 11),
        ("5. Как изменить период продаж и период заказа", True, 13),
        ("Откройте «01_Настройки» → ячейки B5 (дни анализа продаж) и B6 (дни горизонта заказа).", False, 11),
        ("После изменения Excel пересчитает среднедневные, прогноз и заказ на листе «03_Расчёт_заказа».", False, 11),
        ("", False, 11),
        ("6. Как менять коэффициенты", True, 13),
        ("Глобально: «01_Настройки» — Коэф. заказа, Повышающий, Понижающий, множители ABC.", False, 11),
        ("По строке: колонка «Коэф. строки» на листе расчёта (например 1.2 для акции на конкретный товар).", False, 11),
        ("", False, 11),
        ("7. Светофор и статусы", True, 13),
        ("🔴 Критический дефицит — остаток на исходе, срочный заказ.", False, 11),
        ("🟠 Риск OOS — покрытие меньше порога риска.", False, 11),
        ("🟡 Избыточный запас — слишком большое покрытие.", False, 11),
        ("🟣 Неликвид — нет продаж при наличии остатка.", False, 11),
        ("🔵 Приоритет A — важная позиция к заказу.", False, 11),
        ("🟢 Норма / рост — плановый заказ.", False, 11),
        ("⚪ Запас достаточен — заказ не требуется.", False, 11),
        ("", False, 11),
        ("8. Логика формул (кратко)", True, 13),
        ("Среднедневные = Продажи / ДниПродаж", False, 11),
        ("Прогноз = Среднедневные × ДниЗаказа × Тренд × КоэфЗаказа × Uplift × Downlift × МножABC × КоэфСтроки", False, 11),
        ("Страховой = Среднедневные × ДниSafety(ABC) × Тренд", False, 11),
        ("Потребность = Прогноз + Страховой", False, 11),
        ("Рекомендуемый заказ = MAX(расчётный; MAX(0; МинОстаток - Остаток))", False, 11),
        ("Минимальный остаток по умолчанию — 24 шт на каждый SKU (параметр на «01_Настройки»).", False, 11),
        ("Для неликвидов без продаж расчётный заказ обнуляется, но докупка до минимального остатка сохраняется.", False, 11),
        ("", False, 11),
        ("9. Советы", True, 13),
        ("Сначала отфильтруйте 🔴 и 🟠, затем класс A, затем остальное.", False, 11),
        ("Длинные наименования переносятся — увеличьте высоту строки при необходимости.", False, 11),
        ("Перед заказом в 1С сверьте единицы измерения и кратность упаковки вручную.", False, 11),
        ]
    )
    ws.column_dimensions["A"].width = 120
    for i, (text, bold, size) in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=text)
        cell.font = Font(name="Calibri", bold=bold, size=size, color="1F4E79" if bold else "333333")
        cell.alignment = ALIGN_WRAP
        ws.row_dimensions[i].height = 18 if not bold else 22
    ws.freeze_panes = "A2"


def _build_settings(wb: Workbook, meta: Dict[str, Any]) -> None:
    ws = _ws(wb, "01_Настройки")
    ws["A1"] = "НАСТРОЙКИ РАСЧЁТА"
    ws["A1"].font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:C1")

    ws["A2"] = "Жёлтые ячейки — редактируемые. От них зависят формулы на листе «03_Расчёт_заказа»."
    ws.merge_cells("A2:C2")

    rows = [
        (4, "Параметр", "Значение", "Подсказка"),
        (5, "Дни анализа продаж", meta.get("period_days", SETTINGS["default_sales_period_days"]),
         "Делитель для среднедневных продаж"),
        (6, "Дни периода заказа", meta.get("order_period_days", SETTINGS["default_order_period_days"]),
         "Горизонт покрытия / прогноз"),
        (7, "Коэффициент заказа", meta.get("order_coefficient", SETTINGS["order_coefficient"]),
         "Глобальный множитель заказа"),
        (8, "Повышающий коэффициент", meta.get("uplift_coefficient", SETTINGS["uplift_coefficient"]),
         "Акции, сезон, рост трафика"),
        (9, "Понижающий коэффициент", meta.get("downlift_coefficient", SETTINGS["downlift_coefficient"]),
         "Спад, окончание сезона"),
        (10, "Safety stock дни (A)", SETTINGS["safety_stock_days_a"], "Страховой запас для A"),
        (11, "Safety stock дни (B)", SETTINGS["safety_stock_days_b"], "Страховой запас для B"),
        (12, "Safety stock дни (C)", SETTINGS["safety_stock_days_c"], "Страховой запас для C"),
        (13, "Множитель заказа ABC-A", SETTINGS["abc_order_mult_a"], "Приоритет защиты A от OOS"),
        (14, "Множитель заказа ABC-B", SETTINGS["abc_order_mult_b"], "База для B"),
        (15, "Множитель заказа ABC-C", SETTINGS["abc_order_mult_c"], "Осторожный заказ C"),
        (16, "Порог риска OOS, дни", SETTINGS["oos_risk_cover_days"], "Покрытие ниже = риск"),
        (17, "Критический OOS, дни", SETTINGS["critical_oos_cover_days"], "Покрытие ниже = критично"),
        (18, "Порог избытка, дни", SETTINGS["overstock_cover_days"], "Покрытие выше = избыток"),
        (19, "Дни без продаж → неликвид", SETTINGS["dead_stock_days_no_sales"], "Критерий неликвида"),
        (20, "Минимальный остаток, шт", SETTINGS.get("min_stock_target", 24),
         "Докупка заказа до этого уровня по каждому SKU (даже при нулевом остатке / низких продажах)"),
        (21, "Дата начала периода", str(meta.get("date_from", ""))[:10], "Информативно"),
        (22, "Дата окончания периода", str(meta.get("date_to", ""))[:10], "Информативно"),
    ]

    for r, a, b, c in rows:
        ws.cell(row=r, column=1, value=a).border = THIN
        cell_b = ws.cell(row=r, column=2, value=b)
        cell_b.border = THIN
        ws.cell(row=r, column=3, value=c).border = THIN
        if r == 4:
            for col in range(1, 4):
                ws.cell(row=r, column=col).fill = FILL_HEADER
                ws.cell(row=r, column=col).font = FONT_HEADER
        elif r <= 20:
            cell_b.fill = FILL_EDIT
            cell_b.comment = Comment(c, "Система", width=200, height=50)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 45

    ws["A24"] = "Справка по формулам"
    ws["A24"].font = Font(bold=True, size=12, color="1F4E79")
    ws["A25"] = "Прогноз = (Продажи/B5)*B6*Тренд*B7*B8*B9*МножABC*КоэфСтроки"
    ws["A26"] = "Заказ = MAX(расчётный; MAX(0; B20 - Остаток)) - докупка до минимального остатка"


def _build_dashboard(wb: Workbook, df: pd.DataFrame, meta: Dict[str, Any]) -> None:
    ws = _ws(wb, "02_Дашборд")
    ws["A1"] = "ИТОГОВАЯ АНАЛИТИКА ЗАКУПОК"
    ws["A1"].font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:D1")

    cards = [
        (3, "Позиций в расчёте", meta.get("items_count", len(df))),
        (4, "Строк к заказу", meta.get("order_lines", 0)),
        (5, "Суммарный рекомендуемый заказ, ед.", round(meta.get("order_qty_total", 0), 2)),
        (6, "Рисков OOS", meta.get("oos_count", 0)),
        (7, "Неликвидов", meta.get("dead_count", 0)),
        (8, "Избыточных запасов", meta.get("overstock_count", 0)),
        (9, "Класс A / B / C", f"{meta.get('abc_a_count', 0)} / {meta.get('abc_b_count', 0)} / {meta.get('abc_c_count', 0)}"),
        (10, "Период продаж, дни", meta.get("period_days", "")),
        (11, "Горизонт заказа, дни", meta.get("order_period_days", "")),
        (12, "Продажи за период, ед.", round(meta.get("total_sales_qty", 0), 2)),
        (13, "Текущий остаток, ед.", round(meta.get("total_stock", 0), 2)),
    ]
    ws["A2"] = "Показатель"
    ws["B2"] = "Значение"
    style_header_row(ws, 2, 2)
    for r, title, val in cards:
        ws.cell(row=r, column=1, value=title).border = THIN
        cell = ws.cell(row=r, column=2, value=val)
        cell.border = THIN
        cell.fill = FILL_FORMULA

    # Топ приоритетов
    ws["A15"] = "ТОП приоритетных позиций (критично / OOS / A)"
    ws["A15"].font = Font(bold=True, size=12, color="1F4E79")
    top = df.sort_values("priority").head(15)
    headers = ["Артикул", "Наименование", "ABC", "Остаток", "Заказ", "Статус", "Светофор"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=16, column=i, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN
    for r_i, (_, row) in enumerate(top.iterrows(), start=17):
        vals = [
            row["sku"],
            row["name"],
            row["abc_class"],
            row["stock"],
            row["recommended_order"],
            row["status"],
            row["traffic_light"],
        ]
        for c_i, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=v if not isinstance(v, float) else round(v, 3))
            cell.border = THIN
            if c_i == 2:
                cell.alignment = ALIGN_WRAP

    # ABC counts for chart
    ws["F3"] = "ABC"
    ws["G3"] = "Кол-во"
    ws["F4"] = "A"
    ws["G4"] = meta.get("abc_a_count", 0)
    ws["F5"] = "B"
    ws["G5"] = meta.get("abc_b_count", 0)
    ws["F6"] = "C"
    ws["G6"] = meta.get("abc_c_count", 0)
    chart = BarChart()
    chart.title = "Распределение ABC"
    chart.dataLabels = None
    data = Reference(ws, min_col=7, min_row=3, max_row=6)
    cats = Reference(ws, min_col=6, min_row=4, max_row=6)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 12
    chart.height = 8
    ws.add_chart(chart, "F8")

    autosize_columns(ws, name_col=2)
    ws.freeze_panes = "A3"


def _abc_mult_formula(row: int) -> str:
    """Множитель ABC из настроек."""
    return (
        f'IF(H{row}="A",\'01_Настройки\'!$B$13,'
        f'IF(H{row}="B",\'01_Настройки\'!$B$14,\'01_Настройки\'!$B$15))'
    )


def _safety_days_formula(row: int) -> str:
    return (
        f'IF(H{row}="A",\'01_Настройки\'!$B$10,'
        f'IF(H{row}="B",\'01_Настройки\'!$B$11,\'01_Настройки\'!$B$12))'
    )


def _status_formula(row: int) -> str:
    """Статус через формулы Excel — пересчитывается при правках."""
    return (
        f'IF(OR(AND(E{row}>0,O{row}<\'01_Настройки\'!$B$17),AND(E{row}>0,D{row}<=0)),"Критический дефицит",'
        f'IF(AND(E{row}>0,O{row}<\'01_Настройки\'!$B$16),"Риск out-of-stock",'
        f'IF(AND(E{row}<=0,D{row}>0),"Неликвид / зависший",'
        f'IF(AND(E{row}>0,O{row}>\'01_Настройки\'!$B$18),"Избыточный запас",'
        f'IF(AND(N{row}>0,H{row}="A"),"Приоритетный заказ (A)",'
        f'IF(AND(N{row}>0,F{row}>=1.15),"Рост спроса — усилить заказ",'
        f'IF(N{row}>0,"К заказу","Запас достаточен")))))))'
    )


def _risk_formula(row: int) -> str:
    return (
        f'IF(P{row}="Критический дефицит","Высокий риск OOS",'
        f'IF(P{row}="Риск out-of-stock","Средний риск OOS",'
        f'IF(P{row}="Неликвид / зависший","Замороженный капитал",'
        f'IF(P{row}="Избыточный запас","Завышенный stock",'
        f'IF(P{row}="Приоритетный заказ (A)","Контроль наличия",'
        f'IF(P{row}="Рост спроса — усилить заказ","Тренд вверх",'
        f'IF(P{row}="К заказу","Норма","Низкий")))))))'
    )


def _light_formula(row: int) -> str:
    return (
        f'IF(P{row}="Критический дефицит","🔴",'
        f'IF(P{row}="Риск out-of-stock","🟠",'
        f'IF(P{row}="Неликвид / зависший","🟣",'
        f'IF(P{row}="Избыточный запас","🟡",'
        f'IF(P{row}="Приоритетный заказ (A)","🔵",'
        f'IF(OR(P{row}="К заказу",P{row}="Рост спроса — усилить заказ"),"🟢","⚪"))))))'
    )


def _build_main_calc(wb: Workbook, df: pd.DataFrame, meta: Dict[str, Any]) -> None:
    ws = _ws(wb, "03_Расчёт_заказа")
    ws["A1"] = (
        "ОСНОВНОЙ РАСЧЁТ ЗАКАЗА  |  Жёлтые — ввод  |  Зелёные — формулы  |  "
        "Параметры: лист 01_Настройки"
    )
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = THIN

    # Подсказки на заголовках
    tips = {
        4: "Редактируемый остаток",
        5: "Редактируемые продажи за период",
        6: "Тренд: >1 рост, <1 спад",
        7: "Локальный множитель строки",
        8: "Класс ABC (можно скорректировать)",
        14: "Итоговый заказ (формула)",
        20: "Свободный комментарий",
    }
    for col, tip in tips.items():
        ws.cell(row=2, column=col).comment = Comment(tip, "Система", width=180, height=40)

    n = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        r = i + 3  # данные с 3-й строки
        # Значения
        ws.cell(row=r, column=1, value=int(row["row_id"]))
        ws.cell(row=r, column=2, value=safe_str(row["sku"]))
        name_cell = ws.cell(row=r, column=3, value=safe_str(row["name"]))
        name_cell.alignment = ALIGN_WRAP

        stock_cell = ws.cell(row=r, column=4, value=float(row["stock"]))
        sales_cell = ws.cell(row=r, column=5, value=float(row["sales_qty"]))
        trend_cell = ws.cell(row=r, column=6, value=round(float(row["trend_coef"]), 4))
        coef_cell = ws.cell(row=r, column=7, value=1.0)
        abc_cell = ws.cell(row=r, column=8, value=str(row["abc_class"]))

        for cell in (stock_cell, sales_cell, trend_cell, coef_cell, abc_cell, name_cell):
            cell.fill = FILL_EDIT
            cell.border = THIN

        # Формулы
        # I: среднедневные = продажи / дни
        ws.cell(row=r, column=9, value=f"=IF('01_Настройки'!$B$5=0,0,E{r}/'01_Настройки'!$B$5)")
        # J: прогноз
        ws.cell(
            row=r,
            column=10,
            value=(
                f"=I{r}*'01_Настройки'!$B$6*F{r}*'01_Настройки'!$B$7*"
                f"'01_Настройки'!$B$8*'01_Настройки'!$B$9*{_abc_mult_formula(r)}*G{r}"
            ),
        )
        # K: safety
        ws.cell(row=r, column=11, value=f"=I{r}*{_safety_days_formula(r)}*F{r}")
        # L: need
        ws.cell(row=r, column=12, value=f"=J{r}+K{r}")
        # M: need - stock
        ws.cell(row=r, column=13, value=f"=L{r}-D{r}")
        # N: recommended order
        # 1) базовый расчёт с блоком неликвида
        # 2) докупка до минимального остатка B20 (даже при нулевых продажах)
        ws.cell(
            row=r,
            column=14,
            value=(
                f'=MAX(IF(AND(E{r}<=0,F{r}<=1),0,MAX(0,CEILING(M{r},1))),'
                f'MAX(0,CEILING(\'01_Настройки\'!$B$20-D{r},1)))'
            ),
        )
        # O: cover days
        ws.cell(row=r, column=15, value=f"=IF(I{r}>0,D{r}/I{r},IF(D{r}>0,9999,0))")
        # P/Q/R status/risk/light
        ws.cell(row=r, column=16, value=f"={_status_formula(r)}")
        ws.cell(row=r, column=17, value=f"={_risk_formula(r)}")
        ws.cell(row=r, column=18, value=f"={_light_formula(r)}")

        ws.cell(row=r, column=19, value=safe_str(row.get("uom", "")))
        note = ws.cell(row=r, column=20, value="")
        note.fill = FILL_EDIT

        for c in range(1, 21):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN
            if c in (9, 10, 11, 12, 13, 14, 15, 16, 17, 18):
                if cell.fill.fgColor is None or cell.fill.fgColor.rgb == "00000000":
                    cell.fill = FILL_FORMULA
                # не перезаписываем EDIT fill — формульные колонки точно зелёные
            if c >= 9 and c <= 18:
                cell.fill = FILL_FORMULA
            cell.alignment = ALIGN_CENTER if c != 3 else ALIGN_WRAP

        # Высота под длинные названия
        name_len = len(safe_str(row["name"]))
        ws.row_dimensions[r].height = max(18, min(60, 14 + name_len // 40 * 12))

    last_row = n + 2
    if n > 0:
        # Таблица Excel
        table_ref = f"A2:{get_column_letter(len(HEADERS))}{last_row}"
        table = Table(displayName="РасчётЗаказа", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name=SETTINGS["table_style"],
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)
        add_abc_validation(ws, f"H3:H{last_row}")
        add_oos_conditional(ws, "P", 3, last_row)
        add_order_highlight(ws, "N", 3, last_row)
    else:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(HEADERS))}{max(last_row, 2)}"

    ws.freeze_panes = "D3"
    ws.column_dimensions["C"].width = 55
    for col in range(1, len(HEADERS) + 1):
        if col == 3:
            continue
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 14 if col > 3 else 12
    ws.column_dimensions["P"].width = 28
    ws.column_dimensions["Q"].width = 22
    ws.column_dimensions["T"].width = 24


def _build_abc_sheet(wb: Workbook, df: pd.DataFrame) -> None:
    ws = _ws(wb, "04_ABC")
    ws["A1"] = "ABC-АНАЛИЗ (по доле продаж)"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:H1")

    cols = [
        "sku",
        "name",
        "abc_class",
        "abc_rank",
        "sales_qty",
        "sales_amount",
        "abc_share",
        "abc_cum_share",
    ]
    titles = [
        "Артикул",
        "Наименование",
        "ABC",
        "Ранг",
        "Продажи, ед.",
        "Продажи, сумма",
        "Доля",
        "Накопленная доля",
    ]
    for c, t in enumerate(titles, 1):
        cell = ws.cell(row=2, column=c, value=t)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN

    view = df.sort_values("abc_rank") if "abc_rank" in df.columns else df
    for i, (_, row) in enumerate(view.iterrows(), start=3):
        values = [
            row["sku"],
            row["name"],
            row["abc_class"],
            int(row.get("abc_rank", i - 2)),
            float(row["sales_qty"]),
            float(row.get("sales_amount", 0) or 0),
            round(float(row.get("abc_share", 0) or 0), 6),
            round(float(row.get("abc_cum_share", 0) or 0), 6),
        ]
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = THIN
            if c == 2:
                cell.alignment = ALIGN_WRAP
            if c == 3:
                abc = str(v)
                if abc == "A":
                    cell.fill = PatternFill("solid", fgColor=SETTINGS["color_a"])
                elif abc == "B":
                    cell.fill = PatternFill("solid", fgColor=SETTINGS["color_b"])
                else:
                    cell.fill = PatternFill("solid", fgColor=SETTINGS["color_c"])

    last = 2 + len(view)
    if len(view):
        ws.auto_filter.ref = f"A2:H{last}"
    ws.freeze_panes = "A3"
    autosize_columns(ws, name_col=2)


def _build_filtered_sheet(wb: Workbook, subset: pd.DataFrame, title: str, header: str) -> None:
    ws = _ws(wb, title)
    ws["A1"] = header
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:J1")

    titles = [
        "Артикул",
        "Наименование",
        "ABC",
        "Остаток",
        "Продажи",
        "Покрытие, дни",
        "Рек. заказ",
        "Статус",
        "Риск",
        "Светофор",
    ]
    for c, t in enumerate(titles, 1):
        cell = ws.cell(row=2, column=c, value=t)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN

    subset = subset.sort_values("priority") if "priority" in subset.columns else subset
    for i, (_, row) in enumerate(subset.iterrows(), start=3):
        vals = [
            row.get("sku", ""),
            row.get("name", ""),
            row.get("abc_class", ""),
            float(row.get("stock", 0) or 0),
            float(row.get("sales_qty", 0) or 0),
            round(float(row.get("cover_days", 0) or 0), 2),
            float(row.get("recommended_order", 0) or 0),
            row.get("status", ""),
            row.get("risk_level", ""),
            row.get("traffic_light", ""),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = THIN
            if c == 2:
                cell.alignment = ALIGN_WRAP

    last = max(2 + len(subset), 2)
    ws.auto_filter.ref = f"A2:J{last}"
    ws.freeze_panes = "A3"
    if subset.empty:
        ws["A3"] = "Нет позиций в этой категории — отличный знак."
    autosize_columns(ws, name_col=2)


def _build_trends_sheet(wb: Workbook, df: pd.DataFrame) -> None:
    ws = _ws(wb, "08_Тренды")
    ws["A1"] = "ТРЕНДЫ ПРОДАЖ (1-я vs 2-я половина периода)"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:I1")

    titles = [
        "Артикул",
        "Наименование",
        "ABC",
        "Продажи 1 пол.",
        "Продажи 2 пол.",
        "Тренд",
        "Метка",
        "Рек. заказ",
        "Статус",
    ]
    for c, t in enumerate(titles, 1):
        cell = ws.cell(row=2, column=c, value=t)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN

    view = df.sort_values("trend_coef", ascending=False)
    for i, (_, row) in enumerate(view.iterrows(), start=3):
        vals = [
            row["sku"],
            row["name"],
            row["abc_class"],
            float(row.get("sales_h1", 0) or 0),
            float(row.get("sales_h2", 0) or 0),
            round(float(row.get("trend_coef", 1) or 1), 4),
            row.get("trend_label", ""),
            float(row.get("recommended_order", 0) or 0),
            row.get("status", ""),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = THIN
            if c == 2:
                cell.alignment = ALIGN_WRAP
            if c == 7:
                label = str(v)
                if label == "Рост":
                    cell.fill = PatternFill("solid", fgColor=SETTINGS["color_ok"])
                elif label == "Спад":
                    cell.fill = PatternFill("solid", fgColor=SETTINGS["color_warning"])

    last = max(2 + len(view), 2)
    ws.auto_filter.ref = f"A2:I{last}"
    ws.freeze_panes = "A3"
    autosize_columns(ws, name_col=2)


def _build_supplier_order_sheet(wb: Workbook, subset: pd.DataFrame, supplier_name: str) -> None:
    """Лист заказа выбранному поставщику (только при опциональном режиме)."""
    ws = _ws(wb, "09_Заказ_поставщику")
    ws["A1"] = f"ЗАКАЗ ПОСТАВЩИКУ: {supplier_name}"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = FILL_HEADER
    ws.merge_cells("A1:K1")

    ws["A2"] = (
        "Лист формируется только при выбранном поставщике. "
        "Показаны позиции с рекомендуемым заказом > 0."
    )
    ws["A2"].alignment = ALIGN_WRAP
    ws.merge_cells("A2:K2")

    titles = [
        "№",
        "Артикул",
        "Наименование",
        "ABC",
        "Остаток",
        "Продажи",
        "Рек. заказ",
        "Покрытие, дни",
        "Статус",
        "Риск",
        "Ед.изм.",
    ]
    for c, t in enumerate(titles, 1):
        cell = ws.cell(row=3, column=c, value=t)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN

    view = subset.sort_values(
        ["recommended_order", "name"],
        ascending=[False, True],
    ) if not subset.empty and "recommended_order" in subset.columns else subset

    for i, (_, row) in enumerate(view.iterrows(), start=4):
        vals = [
            i - 3,
            row.get("sku", ""),
            row.get("name", ""),
            row.get("abc_class", ""),
            float(row.get("stock", 0) or 0),
            float(row.get("sales_qty", 0) or 0),
            float(row.get("recommended_order", 0) or 0),
            round(float(row.get("cover_days", 0) or 0), 2),
            row.get("status", ""),
            row.get("risk_level", ""),
            row.get("uom", ""),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = THIN
            if c == 3:
                cell.alignment = ALIGN_WRAP
            if c == 7:
                cell.fill = FILL_EDIT

    last = max(3 + len(view), 3)
    ws.auto_filter.ref = f"A3:K{last}"
    ws.freeze_panes = "A4"
    if view.empty:
        ws["A4"] = "По выбранному поставщику нет позиций с рекомендуемым заказом > 0."
    autosize_columns(ws, name_col=3)
