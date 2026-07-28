# Changelog

## Второй релиз

### Добавлено
- Новый слой чтения и нормализации Excel: `data/excel_parser.py`.
- Функции: `normalize_columns`, `detect_sheet`, `detect_header_row`, `parse_sales_file`, `parse_stock_file`, `map_input_columns`, `build_canonical_sales_df`, `build_canonical_stock_df`.
- Расширенная техническая диагностика парсинга (лист, строка заголовка, распознанные колонки, mapping, ключи объединения).

### Изменено
- `data/loaders.py` переведён на новый parser-слой без изменения публичного интерфейса `load_stock_file/load_sales_file`.
- Обновлён стартовый экран `streamlit_app.py`: пометка "Второй релиз", описание новых возможностей загрузки, актуализированная инструкция.
- Обновлён `README.txt` под новый механизм импорта 1С.

### Не изменялось намеренно
- Бизнес-логика расчётов.
- Формулы автозаказа.
- Алгоритмы ABC/прогноза/рисков.
- Архитектура расчётного pipeline.
