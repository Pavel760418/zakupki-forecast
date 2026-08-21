"""Канонические имена магазинов / складов / подразделений 1С."""

from __future__ import annotations

import re

from utils.helpers import safe_str

NETWORK_STORE_LABEL = "Сводно по сети"
UNKNOWN_STORE_LABEL = "Без магазина"

# Центральный склад (не розничная точка) — в отчётах 1С часто единственный в остатках
CENTRAL_WAREHOUSE_KEYS = frozenset(
    {
        "склад основной1",
        "склад основной",
        "сикса",
    }
)

# Короткие имена сети СИКС → канон (как в остатках)
STORE_CANON = {
    "адлер": "Адлер",
    "гагаринский": "Гагаринский",
    "ижевск": "Ижевск",
    "ленинградский": "Ленинградский",
    "орджоникидзе 11": "Орджоникидзе 11",
    "орджоникидзе": "Орджоникидзе 11",
    "перефасовка": "Перефасовка",
    "склад основной1": "Склад основной1",
    "склад основной": "Склад основной1",
    "сикса": "Склад основной1",
    "смоленск галактика": "Смоленск Галактика",
    "сочи приморская": "Сочи Приморская",
    "сочи": "Сочи Приморская",
    "флагман": "Флагман",
}

_PREFIX_RE = re.compile(
    r"^(обособленное подразделение|основное подразделение)\s+",
    re.IGNORECASE,
)


def _norm_key(value: object) -> str:
    text = safe_str(value).strip().casefold().replace("ё", "е")
    text = _PREFIX_RE.sub("", text)
    return re.sub(r"\s+", " ", text).strip()


def canon_store_name(value: object) -> str:
    """Сводит «Сочи» / «Обособленное подразделение Флагман» к одному имени."""
    raw = safe_str(value).strip()
    if not raw:
        return ""
    key = _norm_key(raw)
    if not key or key in {"итого", "всего"}:
        return ""
    if key in STORE_CANON:
        return STORE_CANON[key]
    for alias, canon in STORE_CANON.items():
        if alias in key:
            return canon
    cleaned = _PREFIX_RE.sub("", raw).strip()
    return cleaned or raw


def store_key(value: object) -> str:
    name = canon_store_name(value)
    return name.casefold() if name else ""


def is_central_warehouse(value: object) -> bool:
    """Центральный склад сети (не Адлер/Флагман/Сочи…)."""
    key = store_key(value) or _norm_key(value)
    return bool(key) and key in CENTRAL_WAREHOUSE_KEYS


def has_store_dimension(frame) -> bool:
    """Есть ли в канонической таблице непустые магазины/склады."""
    if frame is None or getattr(frame, "empty", True):
        return False
    if "store" not in frame.columns and "warehouse" not in frame.columns:
        return False
    if "store" in frame.columns:
        values = frame["store"].fillna("").map(lambda x: safe_str(x).strip())
    else:
        values = frame["warehouse"].fillna("").map(lambda x: safe_str(canon_store_name(x)).strip())
    values = values[~values.isin({"", NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL})]
    return bool(len(values))


def has_retail_store_dimension(frame) -> bool:
    """Есть ли розничные точки (без центрального склада и служебных меток)."""
    if frame is None or getattr(frame, "empty", True):
        return False
    col = "store" if "store" in frame.columns else ("warehouse" if "warehouse" in frame.columns else "")
    if not col:
        return False
    values = frame[col].fillna("").map(lambda x: safe_str(canon_store_name(x) if col == "warehouse" else x).strip())
    values = values[~values.isin({"", NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL})]
    values = values[~values.map(is_central_warehouse)]
    return bool(len(values))
