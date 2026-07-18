"""Пошаговый GUI-мастер загрузки файлов и запуска расчёта."""

from __future__ import annotations

import logging
import threading
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

import customtkinter as ctk
import pandas as pd

from calculations.pipeline import run_calculations
from config.settings import SETTINGS
from data.loaders import detect_sales_date_range, load_sales_file, load_stock_file
from excel.workbook_builder import build_workbook
from utils.helpers import ensure_output_dir

logger = logging.getLogger("zakupki_forecast.gui")

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class WizardApp(ctk.CTk):
    """Пошаговый сценарий: остатки → продажи → период → расчёт."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Прогноз заказа — 1С Корп Общепит")
        self.geometry("720x520")
        self.minsize(640, 480)

        self.stock_path: Optional[Path] = None
        self.sales_path: Optional[Path] = None
        self.stock_df: Optional[pd.DataFrame] = None
        self.sales_df: Optional[pd.DataFrame] = None
        self.step = 1

        self._build_ui()
        self._show_step(1)

    def _build_ui(self) -> None:
        self.header = ctk.CTkLabel(
            self,
            text="Модуль прогноза закупок",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        self.header.pack(pady=(18, 4))

        self.subtitle = ctk.CTkLabel(
            self,
            text="Пошаговая загрузка данных из 1С → расчёт → итоговый Excel",
            font=ctk.CTkFont(size=13),
        )
        self.subtitle.pack(pady=(0, 10))

        self.progress = ctk.CTkProgressBar(self, width=560)
        self.progress.pack(pady=6)
        self.progress.set(0.25)

        self.step_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self.step_label.pack(pady=8)

        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=24, pady=8)

        self.status = ctk.CTkLabel(self, text="", text_color="#555555")
        self.status.pack(pady=4)

        self.nav = ctk.CTkFrame(self, fg_color="transparent")
        self.nav.pack(fill="x", padx=24, pady=16)

        self.btn_back = ctk.CTkButton(self.nav, text="← Назад", width=120, command=self._back)
        self.btn_back.pack(side="left")

        self.btn_next = ctk.CTkButton(self.nav, text="Далее →", width=160, command=self._next)
        self.btn_next.pack(side="right")

    def _clear_body(self) -> None:
        for w in self.body.winfo_children():
            w.destroy()

    def _show_step(self, step: int) -> None:
        self.step = step
        self._clear_body()
        self.btn_back.configure(state="normal" if step > 1 else "disabled")
        self.progress.set(step / 4)

        if step == 1:
            self._step_stock()
        elif step == 2:
            self._step_sales()
        elif step == 3:
            self._step_period()
        else:
            self._step_run()

    def _step_stock(self) -> None:
        self.step_label.configure(text="Шаг 1 из 4 — Загрузите файл остатков")
        ctk.CTkLabel(
            self.body,
            text="Выберите Excel-файл остатков из 1С Корп Общепит.\n"
            "Ожидаемые колонки: Артикул/Код, Наименование, Остаток/Количество.",
            justify="left",
        ).pack(anchor="w", pady=8)

        self.stock_var = ctk.StringVar(value=str(self.stock_path) if self.stock_path else "Файл не выбран")
        ctk.CTkLabel(self.body, textvariable=self.stock_var, wraplength=620).pack(anchor="w", pady=6)
        ctk.CTkButton(self.body, text="Выбрать файл остатков…", command=self._pick_stock).pack(anchor="w", pady=8)
        self.btn_next.configure(text="Далее →", state="normal")

    def _step_sales(self) -> None:
        self.step_label.configure(text="Шаг 2 из 4 — Загрузите файл продаж")
        ctk.CTkLabel(
            self.body,
            text="Выберите Excel-файл продаж из 1С.\n"
            "Ожидаемые колонки: Артикул/Код, Наименование, Дата, Количество.",
            justify="left",
        ).pack(anchor="w", pady=8)

        self.sales_var = ctk.StringVar(value=str(self.sales_path) if self.sales_path else "Файл не выбран")
        ctk.CTkLabel(self.body, textvariable=self.sales_var, wraplength=620).pack(anchor="w", pady=6)
        ctk.CTkButton(self.body, text="Выбрать файл продаж…", command=self._pick_sales).pack(anchor="w", pady=8)
        self.btn_next.configure(text="Далее →", state="normal")

    def _step_period(self) -> None:
        self.step_label.configure(text="Шаг 3 из 4 — Укажите период расчёта")

        dmin, dmax = detect_sales_date_range(self.sales_df)
        hint = f"В файле продаж найден период: {dmin.date()} — {dmax.date()}"
        ctk.CTkLabel(self.body, text=hint).pack(anchor="w", pady=6)

        default_days = SETTINGS["default_sales_period_days"]
        start_default = max(dmin, dmax - timedelta(days=default_days - 1))

        form = ctk.CTkFrame(self.body)
        form.pack(fill="x", pady=10)

        ctk.CTkLabel(form, text="Дата начала (ДД.ММ.ГГГГ):").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.date_from_var = ctk.StringVar(value=start_default.strftime("%d.%m.%Y"))
        ctk.CTkEntry(form, textvariable=self.date_from_var, width=160).grid(row=0, column=1, padx=6, pady=6)

        ctk.CTkLabel(form, text="Дата окончания (ДД.ММ.ГГГГ):").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.date_to_var = ctk.StringVar(value=dmax.strftime("%d.%m.%Y"))
        ctk.CTkEntry(form, textvariable=self.date_to_var, width=160).grid(row=1, column=1, padx=6, pady=6)

        ctk.CTkLabel(form, text="Горизонт заказа, дни:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.order_days_var = ctk.StringVar(value=str(SETTINGS["default_order_period_days"]))
        ctk.CTkEntry(form, textvariable=self.order_days_var, width=160).grid(row=2, column=1, padx=6, pady=6)

        ctk.CTkLabel(form, text="Коэффициент заказа:").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.order_coef_var = ctk.StringVar(value=str(SETTINGS["order_coefficient"]))
        ctk.CTkEntry(form, textvariable=self.order_coef_var, width=160).grid(row=3, column=1, padx=6, pady=6)

        ctk.CTkLabel(
            self.body,
            text="Период должен соответствовать (или пересекаться с) датами в загруженном файле продаж.",
            text_color="#666666",
        ).pack(anchor="w", pady=8)

        self.btn_next.configure(text="Далее →", state="normal")

    def _step_run(self) -> None:
        self.step_label.configure(text="Шаг 4 из 4 — Запуск обработки")
        summary = (
            f"Остатки: {self.stock_path.name if self.stock_path else '—'}\n"
            f"Продажи: {self.sales_path.name if self.sales_path else '—'}\n"
            f"Период: {self.date_from_var.get()} — {self.date_to_var.get()}\n"
            f"Горизонт заказа: {self.order_days_var.get()} дн."
        )
        ctk.CTkLabel(self.body, text=summary, justify="left").pack(anchor="w", pady=8)
        ctk.CTkLabel(
            self.body,
            text="Нажмите «Сформировать Excel». Итоговый файл появится в папке output.",
        ).pack(anchor="w", pady=6)

        self.run_btn = ctk.CTkButton(
            self.body,
            text="Сформировать Excel",
            width=220,
            height=40,
            command=self._run_pipeline,
        )
        self.run_btn.pack(anchor="w", pady=16)
        self.btn_next.configure(text="Готово", state="disabled")

    def _pick_stock(self) -> None:
        path = filedialog.askopenfilename(
            title="Файл остатков",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        try:
            self.stock_df = load_stock_file(path)
            self.stock_path = Path(path)
            self.stock_var.set(f"{self.stock_path.name}  ({len(self.stock_df)} позиций)")
            self.status.configure(text="Остатки загружены успешно")
        except Exception as exc:
            logger.exception("Ошибка загрузки остатков")
            messagebox.showerror("Ошибка остатков", str(exc))

    def _pick_sales(self) -> None:
        path = filedialog.askopenfilename(
            title="Файл продаж",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        try:
            self.sales_df = load_sales_file(path)
            self.sales_path = Path(path)
            dmin, dmax = detect_sales_date_range(self.sales_df)
            self.sales_var.set(
                f"{self.sales_path.name}  ({len(self.sales_df)} строк, {dmin.date()} — {dmax.date()})"
            )
            self.status.configure(text="Продажи загружены успешно")
        except Exception as exc:
            logger.exception("Ошибка загрузки продаж")
            messagebox.showerror("Ошибка продаж", str(exc))

    def _back(self) -> None:
        if self.step > 1:
            self._show_step(self.step - 1)

    def _next(self) -> None:
        if self.step == 1:
            if self.stock_df is None:
                messagebox.showwarning("Нужен файл", "Сначала загрузите файл остатков.")
                return
            self._show_step(2)
        elif self.step == 2:
            if self.sales_df is None:
                messagebox.showwarning("Нужен файл", "Сначала загрузите файл продаж.")
                return
            self._show_step(3)
        elif self.step == 3:
            try:
                int(self.order_days_var.get())
                float(self.order_coef_var.get().replace(",", "."))
                datetime.strptime(self.date_from_var.get().strip(), "%d.%m.%Y")
                datetime.strptime(self.date_to_var.get().strip(), "%d.%m.%Y")
            except Exception:
                messagebox.showerror(
                    "Период",
                    "Проверьте даты (ДД.ММ.ГГГГ), горизонт заказа (целое) и коэффициент.",
                )
                return
            self._show_step(4)

    def _run_pipeline(self) -> None:
        self.run_btn.configure(state="disabled", text="Идёт расчёт…")
        self.status.configure(text="Обработка данных…")

        def worker() -> None:
            try:
                ensure_output_dir()
                result_df, meta = run_calculations(
                    self.stock_df,
                    self.sales_df,
                    self.date_from_var.get().strip(),
                    self.date_to_var.get().strip(),
                    order_period_days=int(self.order_days_var.get()),
                    order_coefficient=float(self.order_coef_var.get().replace(",", ".")),
                )
                out = build_workbook(result_df, meta)
                self.after(0, lambda: self._on_success(out, meta))
            except Exception as exc:
                msg = str(exc)
                logger.exception("Ошибка расчёта")
                self.after(0, lambda: self._on_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, path: Path, meta: dict) -> None:
        self.run_btn.configure(state="normal", text="Сформировать Excel")
        self.status.configure(text=f"Готово: {path}")
        messagebox.showinfo(
            "Успех",
            "Итоговый Excel сформирован.\n\n"
            f"Файл:\n{path}\n\n"
            f"Позиций: {meta.get('items_count')}\n"
            f"К заказу: {meta.get('order_lines')}\n"
            f"Рисков OOS: {meta.get('oos_count')}",
        )
        try:
            import os

            os.startfile(path)
        except Exception:
            pass

    def _on_error(self, message: str) -> None:
        self.run_btn.configure(state="normal", text="Сформировать Excel")
        self.status.configure(text="Ошибка расчёта")
        messagebox.showerror("Ошибка", message)


def run_app() -> None:
    app = WizardApp()
    app.mainloop()
