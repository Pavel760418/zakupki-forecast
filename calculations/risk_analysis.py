"""Риски OOS, избыточные запасы, статусы позиций."""

from __future__ import annotations

import pandas as pd

from config.settings import SETTINGS


def apply_risk_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """Присваивает риски, светофор и текстовый статус закупщику."""
    out = df.copy()
    oos_days = SETTINGS["oos_risk_cover_days"]
    crit_days = SETTINGS["critical_oos_cover_days"]
    over_days = SETTINGS["overstock_cover_days"]

    statuses = []
    risks = []
    lights = []
    oos_flags = []
    over_flags = []

    for _, row in out.iterrows():
        cover = float(row.get("cover_days", 0) or 0)
        stock = float(row.get("stock", 0) or 0)
        sales = float(row.get("sales_qty", 0) or 0)
        abc = str(row.get("abc_class", "C"))
        order = float(row.get("recommended_order", 0) or 0)
        dead = bool(row.get("is_dead_stock", False))
        trend = str(row.get("trend_label", ""))

        is_oos_risk = (sales > 0) and (cover < oos_days)
        is_critical = (sales > 0) and (cover < crit_days or stock <= 0)
        is_over = (sales > 0) and (cover > over_days)
        # Без продаж при остатке — неликвид (не избыток в классическом смысле)
        if sales <= 0 and stock > 0:
            is_over = False

        oos_flags.append(is_oos_risk or is_critical)
        over_flags.append(bool(is_over and not is_critical))

        if is_critical:
            status = "Критический дефицит"
            risk = "Высокий риск OOS"
            light = "🔴"
        elif is_oos_risk:
            status = "Риск out-of-stock"
            risk = "Средний риск OOS"
            light = "🟠"
        elif dead:
            status = "Неликвид / зависший"
            risk = "Замороженный капитал"
            light = "🟣"
        elif is_over:
            status = "Избыточный запас"
            risk = "Завышенный stock"
            light = "🟡"
        elif order > 0 and abc == "A":
            status = "Приоритетный заказ (A)"
            risk = "Контроль наличия"
            light = "🔵"
        elif trend == "Рост" and order > 0:
            status = "Рост спроса — усилить заказ"
            risk = "Тренд вверх"
            light = "🟢"
        elif order > 0:
            status = "К заказу"
            risk = "Норма"
            light = "🟢"
        else:
            status = "Запас достаточен"
            risk = "Низкий"
            light = "⚪"

        statuses.append(status)
        risks.append(risk)
        lights.append(light)

    out["status"] = statuses
    out["risk_level"] = risks
    out["traffic_light"] = lights
    out["is_oos_risk"] = oos_flags
    out["is_overstock"] = over_flags

    # Приоритет для сортировки закупщика
    priority_map = {
        "Критический дефицит": 1,
        "Риск out-of-stock": 2,
        "Приоритетный заказ (A)": 3,
        "Рост спроса — усилить заказ": 4,
        "К заказу": 5,
        "Избыточный запас": 6,
        "Неликвид / зависший": 7,
        "Запас достаточен": 8,
    }
    out["priority"] = out["status"].map(lambda s: priority_map.get(s, 99))
    return out
