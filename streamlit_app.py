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
from data.merge import GRAIN_NETWORK, GRAIN_STORE
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


@st.cache_data(show_spinner=False)
def _load_supplier_options() -> tuple[str, ...]:
    """Стабильный tuple опций для selectbox (кэш Streamlit)."""
    try:
        mapping = get_cached_supplier_mapping()
        names = [safe for safe in (str(x).strip() for x in list_suppliers(mapping)) if safe]
        # Уникальные, порядок сохранён
        uniq: list[str] = []
        seen: set[str] = set()
        for name in names:
            if name not in seen and name != SUPPLIER_NONE_LABEL:
                seen.add(name)
                uniq.append(name)
        return (SUPPLIER_NONE_LABEL, *uniq)
    except Exception:
        return (SUPPLIER_NONE_LABEL,)


def _render_supplier_block() -> str | None:
    """
    Опциональный выбор поставщика.

    Важно для Streamlit Cloud:
    - не смешивать index= и key= у selectbox;
    - не менять дерево виджетов между успешным/ошибочным путём;
    - выбирать по индексу (имена поставщиков содержат кавычки и ломают DOM).
    """
    st.subheader("4. Поставщик (опционально)")
    st.caption(
        "По умолчанию выполняется общий расчёт. Выбор поставщика не обязателен "
        "и фильтрует SKU по файлу привязки. Отсутствие остатков в файле не блокирует расчёт."
    )

    labels = list(_load_supplier_options())
    if not labels:
        labels = [SUPPLIER_NONE_LABEL]

    st.caption(f"В справочнике привязки: **{max(0, len(labels) - 1)}** поставщиков.")

    # Выбор по индексу — устойчивее для React DOM (имена с «"» внутри).
    idx_key = "supplier_select_idx"
    st.session_state.pop("supplier_select", None)  # старый key с текстом опции
    current_idx = st.session_state.get(idx_key, 0)
    if not isinstance(current_idx, int) or current_idx < 0 or current_idx >= len(labels):
        st.session_state[idx_key] = 0

    selected_idx = st.selectbox(
        "Поставщик для расчёта заказа",
        options=list(range(len(labels))),
        format_func=lambda i: labels[i],
        key=idx_key,
        help="Оставьте «Не выбран / общий расчёт», чтобы посчитать всю номенклатуру.",
    )
    selected_supplier = labels[int(selected_idx)]

    if selected_supplier != SUPPLIER_NONE_LABEL:
        st.info(f"Режим: расчёт по поставщику **{selected_supplier}**")
        return selected_supplier

    st.info("Режим: **все контрагенты** (поставщик будет указан в строке товара)")
    return None


def main() -> None:
    st.title("Автозаказ СИКС — Четвёртый релиз")
    st.caption(
        "Расчёт автозаказа по остаткам и продажам из 1С. "
        "Можно считать сводно по сети или с разбивкой по магазинам; "
        "поставщик на каждой позиции, отдельный лист заявки — всегда."
    )

    with st.expander("📘 Как пользоваться (пошагово)", expanded=True):
        st.markdown(
            """
**Что нового в четвёртом релизе.**
- Расчёт можно сделать **сводно по сети** или **по магазинам / подразделениям** (группировки 1С: Адлер, Флагман, Сочи…).
- На листе **03_Расчёт_заказа** у каждой позиции есть **магазин, поставщик, штрихкод, цена приходная**.
- Лист **09_Заказ_поставщику** есть **всегда**: наименование, количество, штрихкод, цена, сумма.
- При детализации по магазинам добавляется **10_Матрица_заказ** (SKU × точка).
- Выбор поставщика по-прежнему необязателен: без выбора — все контрагенты, с выбором — только его SKU.
- Для каждого SKU действует **минимальный остаток 24 шт** (параметр на `01_Настройки`).
- Расчёт по поставщику выполняется и **без остатков в файле**: отсутствующие SKU берутся из привязки с остатком 0.

**Откат на 3 релиз.** В GitHub ветка `cursor/backup-release-3-e9d7` и тег `release-3`.
В Streamlit Cloud можно переключить приложение на эту ветку.

**Порядок работы:**
1. Загрузите **остатки** из 1С (подойдёт отчёт с группировкой по складу, как «Остатки на 16.08»).
2. Загрузите **продажи** (кросс по датам и подразделениям, как «Продажи 01.08–16.08»).
3. Проверьте **период** — подставится из файла продаж.
4. Задайте горизонт заказа и коэффициент.
5. *(Опционально)* выберите поставщика.
6. Выберите детализацию: **сводно** или **по магазинам**.
7. Нажмите **«Сформировать Excel»**.

**Выбор поставщика — необязателен.**
- Значение по умолчанию: **«Не выбран / общий расчёт»**.
- Без выбора поставщика выполняется стандартный общий расчёт.
- При выборе поставщика: фильтрация по привязке SKU → поставщик.
- Список поставщиков строится **полностью** из файла привязки (включая строки с объединёнными ячейками).
- Если в остатках нет SKU выбранного поставщика, они всё равно попадают в расчёт с остатком 0 и докупаются до минимума.

**Имена точек.** «Сочи» = «Сочи Приморская», «Орджоникидзе» = «Орджоникидзе 11»,
«Обособленное подразделение Флагман» = «Флагман». Если в файле нет магазина — останется сводный режим.

Модуль сам ищет лист и строку заголовка. Если .xls не читается — пересохраните как *.xlsx.
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
- **Рекомендуемый заказ** = МАКС(расчётный заказ; докупка до **мин. остатка 24 шт** на `01_Настройки`); для неликвида расчётный заказ = 0, но докупка до минимума сохраняется
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
`09_Заказ_поставщику` (всегда: заявка с штрихкодом, ценой и суммой),
`10_Матрица_заказ` (если считали по магазинам).

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

    supplier_mode = _render_supplier_block()

    st.subheader("5. Детализация")
    grain_label = st.radio(
        "Как формировать заказ",
        options=["Сводно по сети", "По магазинам / подразделениям"],
        key="grain_select",
        help="Сводно — одна строка на товар по всей сети. По магазинам — отдельная потребность Адлера, Флагмана, Сочи и т.д.",
    )
    grain = GRAIN_STORE if grain_label.startswith("По магазинам") else GRAIN_NETWORK
    if grain == GRAIN_STORE:
        st.info("Детализация: **по магазинам**. Если в файлах нет склада/подразделения, модуль останется сводным.")
    else:
        st.info("Детализация: **сводно по сети**")

    st.subheader("6. Запуск")
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
                allow_empty_sales=bool(supplier_info.get("allow_empty_sales")),
                grain=grain,
            )
            if supplier_info.get("supplier_selected"):
                meta["supplier_name"] = supplier_info.get("supplier_name", "")
                meta["supplier_sku_keys"] = supplier_info.get("sku_keys", 0)
            if meta.get("store_grain_fallback"):
                st.warning(
                    "В загруженных файлах не найдены магазины/подразделения. "
                    "Сформирован сводный расчёт по сети."
                )

            # Имя поставщика передаём через meta (см. build_workbook),
            # без отдельного kwargs — совместимо со всеми версиями сборщика.
            out_path = build_workbook(result_df, meta)

        st.success("Готово! Скачайте файл ниже.")
        summary = (
            f"Позиций: **{meta.get('items_count')}** · "
            f"К заказу: **{meta.get('order_lines')}** · "
            f"Рисков OOS: **{meta.get('oos_count')}**"
        )
        if supplier_mode:
            summary += f" · Поставщик: **{supplier_mode}**"
        if meta.get("grain") == GRAIN_STORE:
            summary += f" · Магазинов: **{meta.get('store_count', 0)}**"
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
