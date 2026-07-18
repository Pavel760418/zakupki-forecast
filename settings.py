"""
Параметры расчёта по умолчанию для закупщика общепита / розницы.

Все ключевые коэффициенты дублируются на лист «Настройки» итогового Excel,
чтобы пользователь мог менять их без перезапуска Python.
"""

from pathlib import Path

# Корень проекта
PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_ROOT / "output"

SETTINGS = {
    # --- Периоды (дни) ---
    "default_sales_period_days": 30,   # период анализа продаж
    "default_order_period_days": 14,   # горизонт заказа / покрытия
    "min_period_days": 1,
    "max_period_days": 366,

    # --- Коэффициенты заказа ---
    "order_coefficient": 1.0,         # базовый коэффициент заказа
    "uplift_coefficient": 1.0,        # повышающий (акции, сезон)
    "downlift_coefficient": 1.0,      # понижающий
    "risk_coefficient_default": 1.0,  # базовый риск

    # --- Safety stock (дни покрытия) по ABC ---
    "safety_stock_days_a": 5,
    "safety_stock_days_b": 3,
    "safety_stock_days_c": 1,

    # --- Множители ABC к прогнозному заказу ---
    "abc_order_mult_a": 1.15,         # A — усиливаем защиту от OOS
    "abc_order_mult_b": 1.00,
    "abc_order_mult_c": 0.85,         # C — осторожнее с заказом

    # --- Пороги ABC (% накопленной доли продаж, Pareto) ---
    "abc_a_threshold": 0.80,
    "abc_b_threshold": 0.95,

    # --- Тренд ---
    "trend_min": 0.70,                # нижняя граница трендового коэффициента
    "trend_max": 1.50,                # верхняя граница
    "trend_neutral": 1.00,

    # --- Неликвиды / избыток / OOS ---
    "dead_stock_days_no_sales": 45,   # нет продаж N дней → кандидат в неликвиды
    "overstock_cover_days": 60,       # покрытие остатком > N дней → избыток
    "oos_risk_cover_days": 7,         # покрытие < N дней → риск OOS
    "critical_oos_cover_days": 3,     # критический риск
    "block_order_dead_stock": True,   # блокировать заказ для неликвидов без тренда

    # --- Ограничения заказа ---
    "min_order_qty": 0,
    "max_order_multiplier": 5.0,      # заказ не больше N * прогноз без остатка

    # --- Округление ---
    "round_order_up": True,

    # --- Excel ---
    "workbook_name_prefix": "Заказ_прогноз",
    "freeze_header": True,
    "table_style": "TableStyleMedium2",

    # --- Цвета (RGB hex без #) ---
    "color_header": "1F4E79",
    "color_header_font": "FFFFFF",
    "color_a": "C6EFCE",
    "color_b": "FFEB9C",
    "color_c": "FFC7CE",
    "color_critical": "FF6B6B",
    "color_warning": "FFD93D",
    "color_ok": "6BCB77",
    "color_info": "4D96FF",
    "color_edit": "FFF2CC",
    "color_formula": "E2EFDA",
    "color_readonly": "F2F2F2",
}
