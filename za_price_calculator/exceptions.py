"""Пользовательские исключения модуля."""


class ZAPriceCalculatorError(Exception):
    """Базовая ошибка модуля."""


class FileLoadError(ZAPriceCalculatorError):
    """Ошибка загрузки/чтения Excel-файла."""


class ValidationError(ZAPriceCalculatorError):
    """Ошибка валидации структуры или содержимого данных."""


class SheetBuildError(ZAPriceCalculatorError):
    """Ошибка при построении итогового листа Excel."""
