"""Вспомогательные функции без зависимости от бизнес-логики."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from config.settings import OUTPUT_DIR


def ensure_output_dir() -> Path:
    """Создаёт каталог output при необходимости."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR


def timestamp_filename(prefix: str, ext: str = "xlsx") -> str:
    """Имя файла с меткой времени."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{stamp}.{ext}"


def safe_str(value, default: str = "") -> str:
    """
    Безопасное преобразование в строку без потери пробелов и спецсимволов.
    Не обрезает длинные наименования.
    """
    if value is None:
        return default
    try:
        import pandas as pd

        if pd.isna(value):
            return default
    except Exception:
        pass
    text = str(value)
    # Убираем только управляющие символы, не трогая пробелы и кириллицу
    return "".join(ch for ch in text if ch == "\t" or (ord(ch) >= 32) or ch in "\n\r")


def safe_float(value, default: float = 0.0) -> float:
    """Безопасное преобразование в float."""
    if value is None:
        return default
    try:
        import pandas as pd

        if pd.isna(value):
            return default
    except Exception:
        pass
    try:
        if isinstance(value, str):
            value = value.replace("\xa0", "").replace(" ", "").replace(",", ".")
        return float(value)
    except (TypeError, ValueError):
        return default


def normalize_text(value) -> str:
    """Нормализация ключа сопоставления (артикул) без искажения отображаемого имени."""
    text = safe_str(value).strip()
    return text.casefold()
