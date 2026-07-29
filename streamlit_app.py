"""
Веб-интерфейс «Автозаказ СИКС» (Streamlit).

Локальный запуск:
    streamlit run streamlit_app.py

Облако (Streamlit Community Cloud):
    Main file path = streamlit_app.py
"""

from __future__ import annotations

import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

import streamlit as st

# Корень проекта в sys.path (и локально, и в облаке)
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from calculations.pipeline import run_calculations
from config.settings import SETTINGS
from data.loaders import detect_sales_date_range, load_sales_file, load_stock_file
from data.supplier_mapping import (
    SUPPLIER_NONE_LABEL,
    SupplierMappingError,
    filter_frames_by_supplier,
    get_cached_supplier_mapping,
    list_suppliers,
)
from excel.workbook_builder import build_workbook
from utils.helpers import ensure_output_dir
from utils.logging_config import setup_logging

setup_logging()
ensure_output_dir()

st.set_page_config(
    page_title="Автозаказ СИКС",
    page_icon="📦",
    layout="centered",
)


def _save_upload(uploaded_file) -> Path:
    """Сохраняет загруженный файл во временную папку и возвращает путь."""
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return Path(tmp.name)


def _load_supplier_options() -> list[str]:
    """Список поставщиков для UI. При ошибке — только общий режим."""
    try:
        suppliers = list_suppliers(get_cached_supplier_mapping())
        return [SUPPLIER_NONE_LABEL, *suppliers]
    except Exception as exc:
        st.warning(
            "Файл привязки SKU к поставщику недоступен. "
            f"Доступен только общий расчёт. Детали: {exc}"
        )
        return [SUPPLIER_NONE_LABEL]


def main() -> None:
    st.title("Автозаказ СИКС — Третий релиз")
    st.caption(
        "Расчёт автозаказа по остаткам и продажам из 1С. "
        "Выбор поставщика опционален: по умолчанию — общий расчёт по всей номенклатуре."
    )

    with st.expander("📘 Как пользоваться (пошагово)", expanded=True):
        st.markdown(
            """
**Что нового в третьем релизе.**
- Приложение называется **Автозаказ СИКС**.
- Добавлена **опциональная** логика заказа в разрезе поставщика.
- Привязка SKU → поставщик берётся из встроенного файла `Привязка_SKU_к_контрагенту.xlsx`.
- Можно выбрать поставщика и получить отдельный лист **«09_Заказ_поставщику»** в итоговом Excel.
- Если поставщик **не выбран**, приложение работает как раньше: общий расчёт по всей номенклатуре.

**Что делает модуль.** По выгрузкам из 1С (остатки + продажи) он считает прогноз
спроса и **рекомендуемый заказ** по каждой позиции, распределяет товары по классам
**ABC** и помечает риски дефицита (**OOS**), неликвиды и избыток.

**Порядок работы:**
1. Загрузите **файл остатков** (Excel из 1С): номенклатура/код/артикул, количество (остаток).
2. Загрузите **файл продаж** (Excel из 1С): номенклатура/код/артикул, продажи (количество), даты.
3. Проверьте **период расчёта** — он подставляется автоматически из файла продаж, при необходимости поправьте.
4. Задайте **Горизонт заказа, дни** и **Коэффициент заказа**.
5. *(Опционально)* выберите поставщика в блоке ниже — тогда в расчёт попадут только SKU этого поставщика.
6. Нажмите **«Сформировать Excel»** и скачайте готовый файл.

**Выбор поставщика — необязателен.**
- Значение по умолчанию: **«Не выбран / общий расчёт»**.
- Без выбора поставщика выполняется стандартный общий расчёт.
- При выборе поставщика: фильтрация по привязке SKU → поставщик + лист «09_Заказ_поставщику».

Модуль сам ищет подходящий лист и строку заголовка в выгрузке 1С.
Если файл не читается — пересохраните его как «Книга Excel (*.xlsx)» и загрузите снова.
            """
        )

    with st.expander("🧮 Логика расчёта и формулы", expanded=False):
        st.markdown(
            """
Расчёт идёт по слоям (те же формулы зашиты в готовый Excel и пересчитываются при правках):

- **Среднедневные** = Продажи за период ÷ Дни анализа продаж
- **Тренд** = продажи 2-й половины периода ÷ 1-й половины (ограничен 0.70–1.50); >1.15 — «Рост», <0.85 — «Спад»
- **Прогноз** = Среднедневные × Дни заказа × Тренд × Коэф. заказа × Повышающий × Понижающий × Множитель ABC × Коэф. строки
- **Страховой запас** = Среднедневные × Safety-дни (по классу ABC) × Тренд
- **Потребность** = Прогноз + Страховой запас
- **Рекомендуемый заказ** = МАКС(0; округление вверх(Потребность − Остаток)); для неликвида (нет продаж и тренд ≤ 1) заказ = 0
- **Покрытие, дни** = Остаток ÷ Среднедневные

**ABC** (Парето по доле продаж — по сумме, если есть выручка, иначе по количеству):
**A** — до 80 % накопленной доли, **B** — до 95 %, **C** — остальное и позиции без продаж.

**Светофор и статусы:**
- 🔴 Критический дефицит — покрытие ниже критического порога или остатка нет
- 🟠 Риск out-of-stock — покрытие ниже порога риска
- 🟡 Избыточный запас — покрытие выше порога избытка
- 🟣 Неликвид — нет продаж при наличии остатка
- 🔵 Приоритет A — важная позиция к заказу
- 🟢 Норма / рост — плановый заказ; ⚪ запас достаточен
            """
        )

    with st.expander("📊 Работа с готовым Excel: параметры и моделирование", expanded=False):
        st.markdown(
            """
Готовый файл — это **интерактивная модель**: формулы пересчитываются прямо в Excel,
ничего перезапускать не нужно.

**Листы:** `00_Инструкция`, `01_Настройки`, `02_Дашборд`,
`03_Расчёт_заказа` (основной рабочий), `04_ABC`, `05_Риск_OOS`,
`06_Неликвиды`, `07_Избыток`, `08_Тренды`,
`09_Заказ_поставщику` (только если выбран поставщик).

**Цвет ячейки = правило:**
- 🟨 **жёлтые — можно менять** (ручной ввод);
- 🟩 **зелёные — формулы**, их не трогают (перезапишутся при пересчёте).

**Глобальные параметры — лист `01_Настройки` (жёлтые ячейки):** дни анализа продаж,
дни горизонта заказа, коэффициент заказа, повышающий/понижающий коэффициенты,
страховой запас по A/B/C, множители ABC, пороги риска OOS / критики / избытка / неликвида.
Изменение любой из них мгновенно пересчитывает весь лист `03_Расчёт_заказа`.

**Правки по строке — лист `03_Расчёт_заказа` (жёлтые колонки):** Остаток,
Продажи за период, Тренд, **Коэф. строки** (например 1.2 под акцию на конкретный товар),
класс ABC (выпадающий список A/B/C), Наименование, Комментарий закупщика.

**Моделирование «что если»:** меняйте горизонт заказа или коэффициенты на
`01_Настройки` — и сразу видите новый «Рекомендуемый заказ», покрытие и статусы.
Для сценария по одному товару используйте «Коэф. строки». Новые позиции добавляйте
**внизу** таблицы, копируя формулы из соседней строки.

**Перед заказом в 1С** сверьте единицы измерения и кратность упаковки вручную;
сначала отрабатывайте 🔴 и 🟠, затем класс A.
            """
        )

    st.subheader("1. Файл остатков")
    stock_file = st.file_uploader(
        "Выберите Excel с остатками",
        type=["xlsx", "xlsm", "xls"],
        key="stock",
    )

    st.subheader("2. Файл продаж")
    sales_file = st.file_uploader(
        "Выберите Excel с продажами",
        type=["xlsx", "xlsm", "xls"],
        key="sales",
    )

    date_from_default = date.today() - timedelta(days=SETTINGS["default_sales_period_days"] - 1)
    date_to_default = date.today()

    if sales_file is not None:
        try:
            sales_path_preview = _save_upload(sales_file)
            sales_preview = load_sales_file(sales_path_preview)
            dmin, dmax = detect_sales_date_range(sales_preview)
            date_from_default = dmin.date()
            date_to_default = dmax.date()
            file_stamp = f"{sales_file.name}:{sales_file.size}"
            if st.session_state.get("sales_stamp") != file_stamp:
                st.session_state["sales_stamp"] = file_stamp
                st.session_state["date_from"] = date_from_default
                st.session_state["date_to"] = date_to_default
            st.info(f"В файле продаж найден период: **{dmin.date()} — {dmax.date()}**")
        except Exception as exc:
            st.warning(f"Файл продаж пока не удалось прочитать: {exc}")

    st.subheader("3. Период расчёта")
    col1, col2 = st.columns(2)
    with col1:
        date_from = st.date_input("Дата начала", value=date_from_default, key="date_from")
    with col2:
        date_to = st.date_input("Дата окончания", value=date_to_default, key="date_to")

    col3, col4 = st.columns(2)
    with col3:
        order_days = st.number_input(
            "Горизонт заказа, дни",
            min_value=1,
            max_value=366,
            value=int(SETTINGS["default_order_period_days"]),
            step=1,
        )
    with col4:
        order_coef = st.number_input(
            "Коэффициент заказа",
            min_value=0.1,
            max_value=5.0,
            value=float(SETTINGS["order_coefficient"]),
            step=0.1,
        )

    st.subheader("4. Поставщик (опционально)")
    st.caption(
        "По умолчанию выполняется общий расчёт. Выбор поставщика не обязателен "
        "и фильтрует только SKU, закреплённые за ним в файле привязки."
    )
    supplier_options = _load_supplier_options()
    selected_supplier = st.selectbox(
        "Поставщик для расчёта заказа",
        options=supplier_options,
        index=0,
        key="supplier_select",
        help="Оставьте «Не выбран / общий расчёт», чтобы посчитать всю номенклатуру.",
    )
    supplier_mode = (
        selected_supplier
        if selected_supplier and selected_supplier != SUPPLIER_NONE_LABEL
        else None
    )
    if supplier_mode:
        st.info(f"Режим: расчёт по поставщику **{supplier_mode}**")
    else:
        st.info("Режим: **общий расчёт** (без фильтра по поставщику)")

    st.subheader("5. Запуск")
    run = st.button("Сформировать Excel", type="primary", use_container_width=True)

    if not run:
        return

    if stock_file is None or sales_file is None:
        st.error("Загрузите оба файла: остатки и продажи.")
        return

    if date_to < date_from:
        st.error("Дата окончания раньше даты начала.")
        return

    try:
        with st.spinner("Считаем прогноз и формируем Excel…"):
            stock_path = _save_upload(stock_file)
            sales_path = _save_upload(sales_file)

            stock_df = load_stock_file(stock_path)
            sales_df = load_sales_file(sales_path)

            # Опциональная фильтрация: только при явном выборе поставщика
            stock_df, sales_df, supplier_info = filter_frames_by_supplier(
                stock_df,
                sales_df,
                supplier_mode,
            )

            result_df, meta = run_calculations(
                stock_df,
                sales_df,
                date_from.strftime("%d.%m.%Y"),
                date_to.strftime("%d.%m.%Y"),
                order_period_days=int(order_days),
                order_coefficient=float(order_coef),
            )
            if supplier_info.get("supplier_selected"):
                meta["supplier_name"] = supplier_info.get("supplier_name", "")
                meta["supplier_sku_keys"] = supplier_info.get("sku_keys", 0)

            out_path = build_workbook(
                result_df,
                meta,
                supplier_name=supplier_mode,
            )

        st.success("Готово! Скачайте файл ниже.")
        summary = (
            f"Позиций: **{meta.get('items_count')}** · "
            f"К заказу: **{meta.get('order_lines')}** · "
            f"Рисков OOS: **{meta.get('oos_count')}**"
        )
        if supplier_mode:
            summary += f" · Поставщик: **{supplier_mode}**"
        st.write(summary)

        data = Path(out_path).read_bytes()
        st.download_button(
            label="⬇️ Скачать итоговый Excel",
            data=data,
            file_name=Path(out_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except SupplierMappingError as exc:
        st.error(f"Ошибка режима по поставщику: {exc}")
    except Exception as exc:
        st.error(f"Ошибка расчёта: {exc}")
        st.exception(exc)


if __name__ == "__main__":
    main()
