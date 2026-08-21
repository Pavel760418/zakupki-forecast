"""Извлечение кванта (шт в упаковке/коробе) из наименования номенклатуры.

Правила (жёстко зафиксированы, правки — через OVERRIDES в quantum_library):
- после последней «*» / «×» / «x» — число упаковки (для «60г*18шт*6» берём внешний 6);
- в конце строки «/N», «/N уп», «1/N», «1/N шт».
"""

from __future__ import annotations

import re
from typing import Optional

from utils.helpers import safe_str

# 60г*18шт*6  /  40 гр*40 шт  /  120г*30
_STAR_CHUNK = re.compile(
    r"[*×xх]\s*(\d+)\s*(?:шт|шб|уп|кор|пач|пакет|ed)?\.?",
    re.IGNORECASE,
)
# 160 гр/80 уп.  /  80 гр/24
_SLASH_PACK = re.compile(
    r"/\s*(\d+)\s*(?:уп|шт|шб|кор|пач|пакет)?\.?\s*$",
    re.IGNORECASE,
)
# 1/24  /  1/18шт  /  1/36 пакет
_ONE_SLASH = re.compile(
    r"(?:^|[\s,;])1\s*/\s*(\d+)\s*(?:шт|уп|шб|кор|пач|пакет)?\.?\s*$",
    re.IGNORECASE,
)
# запасной: «... 24 шт» / «... 48шт» в конце без * (слабее приоритет)
_TRAILING_PCS = re.compile(
    r"(?<![*/\d])(\d+)\s*шт\.?\s*$",
    re.IGNORECASE,
)

# Слишком большие числа почти наверняка вес/штрих, не квант
_MAX_REASONABLE = 5000
_MIN_REASONABLE = 2


def parse_quantum_from_name(name: object) -> Optional[int]:
    """Возвращает квант (целое ≥2) или None, если в названии не найден."""
    text = safe_str(name).strip()
    if not text:
        return None

    stars = list(_STAR_CHUNK.finditer(text))
    if stars:
        value = int(stars[-1].group(1))
        if _MIN_REASONABLE <= value <= _MAX_REASONABLE:
            return value

    m_one = _ONE_SLASH.search(text)
    if m_one:
        value = int(m_one.group(1))
        if _MIN_REASONABLE <= value <= _MAX_REASONABLE:
            return value

    m_slash = _SLASH_PACK.search(text)
    if m_slash:
        value = int(m_slash.group(1))
        if _MIN_REASONABLE <= value <= _MAX_REASONABLE:
            return value

    m_trail = _TRAILING_PCS.search(text)
    if m_trail:
        value = int(m_trail.group(1))
        # «40 гр 40 шт» без звёздочки — допускаем; отсекаем «1 шт»
        if _MIN_REASONABLE <= value <= _MAX_REASONABLE:
            return value

    return None


def ceil_to_quantum(qty: float, quantum: int | float | None) -> float:
    """Округляет количество вверх до кратного кванта. qty≤0 → 0."""
    q = float(qty or 0.0)
    if q <= 0:
        return 0.0
    pack = int(quantum or 0)
    if pack <= 1:
        return float(q)
    # целые упаковки
    units = int(q) if abs(q - int(q)) < 1e-9 else q
    import math

    n_packs = int(math.ceil(units / pack - 1e-12))
    return float(n_packs * pack)
