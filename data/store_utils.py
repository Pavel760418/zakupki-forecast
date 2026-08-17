"""Канонические имена магазинов / складов / подразделений 1С."""

from __future__ import annotations

import re

from utils.helpers import safe_str

NETWORK_STORE_LABEL = "Сводно по сети"
UNKNOWN_STORE_LABEL = "Без магазина"

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


def has_store_dimension(frame) -> bool:
    """Есть ли в канонической таблице непустые магазины."""
    if frame is None or getattr(frame, "empty", True):
        return False
    if "store" not in frame.columns:
        return False
    values = frame["store"].fillna("").map(lambda x: safe_str(x).strip())
    values = values[~values.isin({"", NETWORK_STORE_LABEL, UNKNOWN_STORE_LABEL})]
    return bool(len(values))
