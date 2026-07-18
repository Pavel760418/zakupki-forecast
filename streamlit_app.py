"""
Веб-интерфейс для менеджера по закупкам (Streamlit).

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
from excel.workbook_builder import build_workbook
from utils.helpers import ensure_output_dir
from utils.logging_config import setup_logging

setup_logging()
ensure_output_dir()

st.set_page_config(
    page_title="Прогноз заказа",
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


def main() -> None:
    st.title("Прогноз заказа")
    st.caption("Загрузите остатки и продажи из 1С → получите Excel с заказом, ABC и рисками")

    with st.expander("Как пользоваться (кратко)", expanded=False):
        st.markdown(
            """
1. Загрузите **файл остатков** (Excel).
2. Загрузите **файл продаж** (Excel).
3. Укажите **период** и горизонт заказа.
4. Нажмите **Сформировать Excel**.
5. Скачайте готовый файл и работайте в нём (листы «Настройки», «Расчёт_заказа»).

Можно использовать шаблоны из папки `samples/templates`.
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

    st.subheader("4. Запуск")
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

            result_df, meta = run_calculations(
                stock_df,
                sales_df,
                date_from.strftime("%d.%m.%Y"),
                date_to.strftime("%d.%m.%Y"),
                order_period_days=int(order_days),
                order_coefficient=float(order_coef),
            )
            out_path = build_workbook(result_df, meta)

        st.success("Готово! Скачайте файл ниже.")
        st.write(
            f"Позиций: **{meta.get('items_count')}** · "
            f"К заказу: **{meta.get('order_lines')}** · "
            f"Рисков OOS: **{meta.get('oos_count')}**"
        )

        data = Path(out_path).read_bytes()
        st.download_button(
            label="⬇️ Скачать итоговый Excel",
            data=data,
            file_name=Path(out_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(f"Ошибка расчёта: {exc}")
        st.exception(exc)


if __name__ == "__main__":
    main()
