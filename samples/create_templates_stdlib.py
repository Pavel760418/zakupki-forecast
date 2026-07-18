"""
Минимальный генератор xlsx без зависимостей кроме zip/xml stdlib —
используется только если openpyxl недоступен.
Но основной путь: create_templates.py с openpyxl.
Этот скрипт — запасной для среды без pip.
"""
from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parent / "templates"


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_xml(rows: list[list[str]], sheet_name_ignored: str = "") -> str:
    """rows: list of rows, each cell is string. Row 1 = headers."""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheetData>",
    ]
    for r_i, row in enumerate(rows, start=1):
        parts.append(f'<row r="{r_i}">')
        for c_i, val in enumerate(row, start=1):
            ref = f"{col_letter(c_i)}{r_i}"
            text = escape("" if val is None else str(val))
            parts.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'
            )
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for i, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(sheets)}</sheets></workbook>'
    )


def rels_xml(n_sheets: int) -> str:
    rels = []
    for i in range(1, n_sheets + 1):
        rels.append(
            f'<Relationship Id="rId{i}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(rels)}</Relationships>'
    )


def content_types(n_sheets: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    for i in range(1, n_sheets + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{"".join(overrides)}</Types>'
    )


def root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def write_xlsx(path: Path, sheets: dict[str, list[list[str]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    names = list(sheets.keys())
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(names)))
        z.writestr("_rels/.rels", root_rels())
        z.writestr("xl/workbook.xml", workbook_xml(names))
        z.writestr("xl/_rels/workbook.xml.rels", rels_xml(len(names)))
        for i, name in enumerate(names, start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(sheets[name]))


def main() -> None:
    stock_data = [
        ["Артикул", "Наименование", "Остаток", "Ед. изм.", "Склад"],
        ["A-001", "Молоко ультрапастеризованное 3,2% 1л", "12", "л", "Основной"],
        ["A-002", "Хлеб пшеничный нарезка 500г", "40", "шт", "Основной"],
        ["A-003", "Куриное филе охлаждённое 1кг", "5", "кг", "Холодильник"],
    ]
    # пустые строки для ввода (не добавляем текст-подсказки на этот лист —
    # любая строка с текстом в «Артикул» попадёт в загрузку)
    for _ in range(30):
        stock_data.append(["", "", "", "", ""])

    stock_instr = [
        ["ШАБЛОН ОСТАТКОВ — ИНСТРУКЦИЯ"],
        [""],
        ["1. Куда вносить данные"],
        [
            "Только на лист «Данные». Первая строка — заголовки, со второй — товары. "
            "Этот лист «Инструкция» модуль при загрузке полностью игнорирует."
        ],
        [""],
        ["2. Обязательные колонки"],
        ["Артикул — код товара из 1С."],
        ["Наименование — полное название товара."],
        ["Остаток — количество на складе (число)."],
        [""],
        ["3. Необязательные колонки"],
        ["Ед. изм. — единица измерения."],
        ["Склад — склад хранения."],
        [""],
        ["4. Как заполнять"],
        ["Скопируйте данные из выгрузки 1С под заголовки строки 1."],
        ["Длинные наименования не сокращайте."],
        ["Артикул лучше как текст (если начинается с нуля)."],
        ["Пустые строки внизу можно не удалять."],
        [""],
        ["5. Чего нельзя делать"],
        ["Не переименовывайте лист «Данные»."],
        ["Не меняйте заголовки в строке 1."],
        ["Не пишите инструкцию на листе «Данные» над таблицей."],
        ["Не объединяйте ячейки в таблице данных."],
        [""],
        ["6. Как загрузить в модуль"],
        ["Сохраните файл → программа → шаг «Загрузите файл остатков» → выберите этот файл."],
        [""],
        ["7. Примеры"],
        ["Строки с A-001… — примеры. Перед реальной работой замените или удалите их."],
    ]

    sales_data = [
        ["Артикул", "Наименование", "Дата", "Количество", "Сумма"],
        ["A-001", "Молоко ультрапастеризованное 3,2% 1л", "01.06.2026", "2", "200"],
        ["A-001", "Молоко ультрапастеризованное 3,2% 1л", "02.06.2026", "3", "300"],
        ["A-002", "Хлеб пшеничный нарезка 500г", "01.06.2026", "5", "250"],
        ["A-003", "Куриное филе охлаждённое 1кг", "03.06.2026", "1.5", "450"],
    ]
    for _ in range(40):
        sales_data.append(["", "", "", "", ""])

    sales_instr = [
        ["ШАБЛОН ПРОДАЖ — ИНСТРУКЦИЯ"],
        [""],
        ["1. Куда вносить данные"],
        [
            "Только на лист «Данные». Строка 1 — заголовки, со строки 2 — продажи. "
            "Лист «Инструкция» модуль игнорирует."
        ],
        [""],
        ["2. Обязательные колонки"],
        ["Артикул — код товара."],
        ["Наименование — название товара."],
        ["Дата — дата продажи (ДД.ММ.ГГГГ)."],
        ["Количество — сколько продано."],
        [""],
        ["3. Необязательные колонки"],
        ["Сумма — выручка. Если заполнить, ABC по сумме; если пусто — по количеству."],
        [""],
        ["4. Как заполнять"],
        ["Одна строка = один товар в одну дату."],
        ["Один товар в разные дни — несколько строк."],
        ["Период в модуле указывайте по датам из этого файла."],
        [""],
        ["5. Чего нельзя делать"],
        ["Не переименовывайте лист «Данные»."],
        ["Не меняйте заголовки строки 1 без необходимости."],
        ["Не пишите пояснения на листе «Данные» выше таблицы."],
        ["Не оставляйте пустую дату в строках с продажами."],
        [""],
        ["6. Как загрузить в модуль"],
        ["Сохраните файл → шаг «Загрузите файл продаж» → выберите этот файл."],
        [""],
        ["7. Примеры"],
        ["Строки с примерами замените своими данными из 1С перед работой."],
    ]

    # Важно: «Данные» — первый лист
    write_xlsx(
        OUT / "Шаблон_остатки.xlsx",
        {"Данные": stock_data, "Инструкция": stock_instr},
    )
    write_xlsx(
        OUT / "Шаблон_продажи.xlsx",
        {"Данные": sales_data, "Инструкция": sales_instr},
    )
    print(OUT / "Шаблон_остатки.xlsx")
    print(OUT / "Шаблон_продажи.xlsx")


if __name__ == "__main__":
    main()
