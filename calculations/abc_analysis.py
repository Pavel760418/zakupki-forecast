"""ABC-анализ по объёму продаж (или сумме, если есть выручка)."""

from __future__ import annotations

import pandas as pd

from config.settings import SETTINGS


def apply_abc_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """
    Классический Pareto ABC по доле продаж.
    Приоритет метрики: sales_amount > 0 → по сумме, иначе по sales_qty.
    """
    out = df.copy()
    use_amount = out["sales_amount"].fillna(0).sum() > 0
    metric = "sales_amount" if use_amount else "sales_qty"

    out["_abc_base"] = out[metric].clip(lower=0)
    total = out["_abc_base"].sum()
    if total <= 0:
        out["abc_class"] = "C"
        out["abc_share"] = 0.0
        out["abc_cum_share"] = 0.0
        out["abc_rank"] = range(1, len(out) + 1)
        out.drop(columns=["_abc_base"], inplace=True)
        return out

    out = out.sort_values("_abc_base", ascending=False).reset_index(drop=True)
    out["abc_share"] = out["_abc_base"] / total
    out["abc_cum_share"] = out["abc_share"].cumsum()
    out["abc_rank"] = out.index + 1

    a_th = SETTINGS["abc_a_threshold"]
    b_th = SETTINGS["abc_b_threshold"]

    def _cls(cum: float) -> str:
        if cum <= a_th:
            return "A"
        if cum <= b_th:
            return "B"
        return "C"

    out["abc_class"] = out["abc_cum_share"].map(_cls)
    # Товары без продаж — C
    out.loc[out["_abc_base"] <= 0, "abc_class"] = "C"
    out.drop(columns=["_abc_base"], inplace=True)
    return out
