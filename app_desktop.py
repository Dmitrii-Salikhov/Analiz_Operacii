"""
Анализ операционной деятельности → «Операции сводная».
Простой UI на классическом tk (надёжно на macOS Tk 8.5).
"""
from __future__ import annotations

import logging
import os
import subprocess
import sys
import traceback
from datetime import datetime, timedelta
from pathlib import Path

os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")

import openpyxl
import pandas as pd
import tkinter as tk
import yaml
from openpyxl.styles import Alignment, Font
from tkinter import BooleanVar, Menu, StringVar, filedialog, messagebox, ttk

from analyzers.app_log import AppLog
from analyzers.emk_compare import compare_plan_emergency, format_mismatch_report
from analyzers.export_report import export_month_like_summary
from analyzers.file_lock import excel_file_locked
from analyzers.form_4001 import compute_form_4001, form_4001_preview_rows
from analyzers.io_utils import OperationsStore, read_table
from analyzers.surgery import SurgeryAnalyzer, build_summary_tables
from analyzers.summary_writer import (
    MONTH_RU,
    SummaryWriter,
    compute_month_weeks,
    read_sheet_weeks,
)
from analyzers.ui_settings import load_settings, save_settings
from analyzers.updater import (
    apply_update_from_zip,
    check_for_update,
    format_update_notes,
    read_local_version,
    resolve_token,
)
from analyzers.year_template import create_year_summary, suggest_summary_path

try:
    from tkcalendar import Calendar
except ImportError:  # pragma: no cover
    Calendar = None  # type: ignore

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
os.chdir(APP_DIR)
APP_LOG = AppLog(APP_DIR / "analysis.log", max_lines=500)

# Дублируем в стандартный logging (без отдельного file handler — пишет AppLog)
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def _btn(parent, text, command, **pack):
    """Обычная tk.Button — всегда видна на macOS."""
    b = tk.Button(parent, text=text, command=command, padx=8, pady=4)
    b.pack(**pack)
    return b


class DesktopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Сводная операций — Видновская КБ")
        self.root.geometry("1200x780")
        self.root.minsize(900, 600)

        self.config = self.load_config()
        if self.config is None:
            root.destroy()
            return

        self.summary_cfg = self.config.get("summary", {})
        self.df_emk = None
        self.emk_path = None
        self.store = OperationsStore()
        self.last_batch_span = (None, None)
        self.cat_table = None
        self.totals_df = None
        self.weeks = []
        self.last_emk_compare = None

        self.plan_mode = StringVar(value="template")
        self.hide_zeros = BooleanVar(value=False)
        self.filter_enabled = BooleanVar(value=False)
        default_summary = self.summary_cfg.get("default_path", "Операции сводная 2026.xlsx")
        self.summary_path = StringVar(value=str(APP_DIR / default_summary))
        self.dept_var = StringVar(value=self.config["departments"]["main"])
        self.preview_month = StringVar()
        self.status_var = StringVar(value="Готов к работе")
        self.start_date_var = StringVar(value="01.01.2026")
        self.end_date_var = StringVar(value="31.12.2026")
        self.year_var = StringVar(value=str(self.summary_cfg.get("year", 2026)))
        self._preview_clipboard = ""
        self.write_weeks_var = BooleanVar(value=True)
        self.write_form_var = BooleanVar(value=True)

        self._apply_saved_settings()

        self.kpi_ops_var = StringVar(value="—")
        self.kpi_patients_var = StringVar(value="—")
        self.kpi_plan_var = StringVar(value="—")
        self.kpi_emerg_var = StringVar(value="—")
        self.kpi_period_var = StringVar(value="—")
        self.kpi_files_var = StringVar(value="—")
        self.kpi_diff_var = StringVar(value="—")

        self._build_menu()
        self._build_layout()
        self._set_date_widgets_state("normal" if self.filter_enabled.get() else "disabled")
        self._bind_shortcuts()
        self._refresh_sources_list()
        self._load_log_into_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log_message("Приложение готово. Нажмите «Опержурнал(ы)» для загрузки.")
        if self.config.get("updates", {}).get("check_on_startup"):
            self.root.after(800, lambda: self.check_updates(silent=True))

    def load_config(self):
        try:
            with open(APP_DIR / "config.yaml", "r", encoding="utf-8") as f:
                return yaml.safe_load(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить config.yaml:\n{e}")
            logging.critical(f"Config load error: {e}")
            return None

    def _build_menu(self):
        bar = Menu(self.root)
        self.root.config(menu=bar)
        file_m = Menu(bar, tearoff=0)
        file_m.add_command(label="Опержурнал(ы)…", command=self.load_surg)
        file_m.add_command(label="ЭМК…", command=self.load_emk)
        file_m.add_command(label="Сводная…", command=self.choose_summary)
        file_m.add_separator()
        file_m.add_command(label="Записать в Excel…", command=self.update_summary)
        file_m.add_command(label="Открыть Excel", command=self.open_summary_file)
        file_m.add_command(label="Экспорт простого отчёта…", command=self.export_simple)
        file_m.add_command(label="Создать сводную на год…", command=self.create_year_summary_dialog)
        file_m.add_command(label="Экспорт неклассифицированных…", command=self.export_unclassified)
        file_m.add_separator()
        file_m.add_command(label="Очистить", command=self.clear_store)
        file_m.add_command(label="Выход", command=self._on_close)
        bar.add_cascade(label="Файл", menu=file_m)
        help_m = Menu(bar, tearoff=0)
        help_m.add_command(label="Проверить обновления…", command=self.check_updates)
        help_m.add_command(label="О программе", command=self.show_about)
        bar.add_cascade(label="Помощь", menu=help_m)

    def _bind_shortcuts(self):
        self.root.bind("<Command-o>", lambda e: self.load_surg())
        self.root.bind("<Control-o>", lambda e: self.load_surg())
        self.root.bind("<Command-s>", lambda e: self.update_summary())
        self.root.bind("<Control-s>", lambda e: self.update_summary())
        self.root.bind("<Command-c>", self._copy_focused_tree)
        self.root.bind("<Control-c>", self._copy_focused_tree)

    def _build_layout(self):
        # ВАЖНО для macOS Tk: виджеты side=BOTTOM паковать ПЕРВЫМИ,
        # иначе expand-область «съедает» окно и оно выглядит пустым.
        tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor="w").pack(
            fill=tk.X, side=tk.BOTTOM
        )

        # --- кнопки сверху ---
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(fill=tk.X, padx=4, pady=4, side=tk.TOP)
        tk.Label(toolbar, text="Действия:", font=("Helvetica", 12, "bold")).pack(side=tk.LEFT, padx=6)
        for text, cmd in (
            ("Опержурнал(ы)", self.load_surg),
            ("ЭМК", self.load_emk),
            ("Обновить превью", self.run_analysis),
            ("Записать в Excel…", self.update_summary),
            ("Открыть Excel", self.open_summary_file),
            ("Расхождения ЭМК", self.show_emk_diff),
            ("Очистить", self.clear_store),
        ):
            _btn(toolbar, text, cmd, side=tk.LEFT, padx=3, pady=4)

        # --- заголовок / отделение ---
        top = tk.Frame(self.root)
        top.pack(fill=tk.X, padx=8, pady=2)
        tk.Label(top, text="Сводная операционной деятельности", font=("Helvetica", 14, "bold")).pack(side=tk.LEFT)
        tk.Label(top, text="  Отделение:").pack(side=tk.LEFT)
        dept_list = self.config["departments"]["list"]
        self.dept_combo = ttk.Combobox(top, textvariable=self.dept_var, values=dept_list, width=40, state="readonly")
        self.dept_combo.pack(side=tk.LEFT, padx=4)

        # --- KPI ---
        kpi = tk.LabelFrame(self.root, text="Сводка", padx=6, pady=4)
        kpi.pack(fill=tk.X, padx=8, pady=4)
        for i, (title, var) in enumerate(
            (
                ("Операций", self.kpi_ops_var),
                ("Пациентов", self.kpi_patients_var),
                ("План %", self.kpi_plan_var),
                ("Экстр. %", self.kpi_emerg_var),
                ("Период", self.kpi_period_var),
                ("Файлы/ЭМК", self.kpi_files_var),
                ("Расхожд.", self.kpi_diff_var),
            )
        ):
            f = tk.Frame(kpi)
            f.grid(row=0, column=i, padx=8, sticky="w")
            tk.Label(f, text=title, fg="#555").pack(anchor="w")
            tk.Label(f, textvariable=var, font=("Helvetica", 14, "bold")).pack(anchor="w")

        # --- настройки ---
        opts = tk.LabelFrame(self.root, text="Настройки", padx=6, pady=4)
        opts.pack(fill=tk.X, padx=8, pady=2)

        r1 = tk.Frame(opts)
        r1.pack(fill=tk.X, pady=2)
        tk.Checkbutton(
            r1, text="Фильтр дат", variable=self.filter_enabled, command=self.toggle_date_widgets
        ).pack(side=tk.LEFT)
        tk.Label(r1, text="с").pack(side=tk.LEFT, padx=(8, 2))
        self.start_entry = tk.Entry(r1, textvariable=self.start_date_var, width=11, state="disabled")
        self.start_entry.pack(side=tk.LEFT)
        self.start_cal_btn = tk.Button(
            r1, text="…", width=2, command=lambda: self._pick_date(self.start_date_var), state="disabled"
        )
        self.start_cal_btn.pack(side=tk.LEFT, padx=(1, 0))
        tk.Label(r1, text="по").pack(side=tk.LEFT, padx=(6, 2))
        self.end_entry = tk.Entry(r1, textvariable=self.end_date_var, width=11, state="disabled")
        self.end_entry.pack(side=tk.LEFT)
        self.end_cal_btn = tk.Button(
            r1, text="…", width=2, command=lambda: self._pick_date(self.end_date_var), state="disabled"
        )
        self.end_cal_btn.pack(side=tk.LEFT, padx=(1, 0))
        tk.Checkbutton(
            r1, text="Скрыть нулевые", variable=self.hide_zeros, command=self.refresh_preview
        ).pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(r1, text="Месяц:").pack(side=tk.LEFT, padx=(12, 2))
        self.month_combo = ttk.Combobox(r1, textvariable=self.preview_month, width=16, state="readonly")
        self.month_combo.pack(side=tk.LEFT)
        self.month_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())
        tk.Label(r1, text="Год:").pack(side=tk.LEFT, padx=(12, 2))
        self.year_combo = ttk.Combobox(
            r1, textvariable=self.year_var, width=6, values=[str(y) for y in range(2024, 2036)]
        )
        self.year_combo.pack(side=tk.LEFT)
        self.year_combo.bind("<<ComboboxSelected>>", lambda e: self._on_year_changed())
        self.year_combo.bind("<Return>", lambda e: self._on_year_changed())

        r2 = tk.Frame(opts)
        r2.pack(fill=tk.X, pady=2)
        tk.Label(r2, text="План/экстренно:").pack(side=tk.LEFT)
        tk.Radiobutton(
            r2, text="По шаблону", variable=self.plan_mode, value="template", command=self._on_plan_mode_change
        ).pack(side=tk.LEFT)
        tk.Radiobutton(
            r2, text="Сверка с ЭМК", variable=self.plan_mode, value="emk", command=self.on_emk_mode
        ).pack(side=tk.LEFT, padx=8)

        r3 = tk.Frame(opts)
        r3.pack(fill=tk.X, pady=2)
        tk.Label(r3, text="Файл сводной:").pack(side=tk.LEFT)
        tk.Entry(r3, textvariable=self.summary_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        _btn(r3, "Обзор…", self.choose_summary, side=tk.LEFT)

        # --- тело: источники + вкладки ---
        body = tk.Frame(self.root)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        left = tk.LabelFrame(body, text="Источники", padx=4, pady=4)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        # Text с переносом — Listbox не умеет wrap длинных имён
        src_wrap = tk.Frame(left)
        src_wrap.pack(fill=tk.BOTH, expand=True)
        self.sources_text = tk.Text(
            src_wrap, width=36, height=18, wrap=tk.WORD, cursor="arrow", relief=tk.SUNKEN, bd=1
        )
        src_vs = ttk.Scrollbar(src_wrap, orient=tk.VERTICAL, command=self.sources_text.yview)
        self.sources_text.configure(yscrollcommand=src_vs.set)
        self.sources_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        src_vs.pack(side=tk.RIGHT, fill=tk.Y)
        self.sources_text.bind("<Double-Button-1>", self._on_source_dblclick)
        self.sources_text.configure(state=tk.DISABLED)
        self.emk_status = tk.Label(left, text="ЭМК: не загружен", anchor="w", wraplength=260, justify=tk.LEFT)
        self.emk_status.pack(fill=tk.X, pady=4)

        right = tk.Frame(body)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self.tab_preview_cat = tk.Frame(self.notebook)
        self.tab_preview_tot = tk.Frame(self.notebook)
        self.tab_preview_form = tk.Frame(self.notebook)
        self.tab_preview = self.tab_preview_cat
        self.tab_uncl = tk.Frame(self.notebook)
        self.tab_log = tk.Frame(self.notebook)
        self.notebook.add(self.tab_preview_cat, text="Превью: категории")
        self.notebook.add(self.tab_preview_tot, text="Превью: итоги")
        self.notebook.add(self.tab_preview_form, text="Превью: форма 4001")
        self.notebook.add(self.tab_uncl, text="Не классифицировано")
        self.notebook.add(self.tab_log, text="Журнал")

        pbtns = tk.Frame(self.tab_preview_cat)
        pbtns.pack(fill=tk.X, pady=2)
        _btn(pbtns, "Записать в Excel…", self.update_summary, side=tk.LEFT, padx=2)
        _btn(pbtns, "Экспорт простого отчёта", self.export_simple, side=tk.LEFT, padx=2)
        _btn(pbtns, "Копировать превью", self.copy_preview, side=tk.LEFT, padx=2)
        self.preview_info = StringVar(value="Превью: нет данных — загрузите опержурнал")
        tk.Label(self.tab_preview_cat, textvariable=self.preview_info, anchor="w").pack(fill=tk.X, padx=2)

        self.tree_preview_cat = self._make_tree(self.tab_preview_cat)
        self.tree_preview_tot = self._make_tree(self.tab_preview_tot)
        form_btns = tk.Frame(self.tab_preview_form)
        form_btns.pack(fill=tk.X, pady=2)
        _btn(form_btns, "Записать в Excel…", self.update_summary, side=tk.LEFT, padx=2)
        _btn(form_btns, "Копировать превью", self.copy_preview, side=tk.LEFT, padx=2)
        self.tree_form = self._make_tree(self.tab_preview_form)

        uncl_top = tk.Frame(self.tab_uncl)
        uncl_top.pack(fill=tk.X, pady=2)
        _btn(uncl_top, "Экспорт списка…", self.export_unclassified, side=tk.LEFT, padx=2)
        uncl_cols = ("Дата", "КВС", "Код", "Название КСГ", "КСГ", "Услуга")
        uncl_body = tk.Frame(self.tab_uncl)
        uncl_body.pack(fill=tk.BOTH, expand=True)
        self.tree_uncl = ttk.Treeview(uncl_body, columns=uncl_cols, show="headings")
        for c, w in zip(uncl_cols, (90, 90, 110, 220, 120, 320)):
            self.tree_uncl.heading(c, text=c)
            self.tree_uncl.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(uncl_body, orient=tk.VERTICAL, command=self.tree_uncl.yview)
        self.tree_uncl.configure(yscrollcommand=vs.set)
        self.tree_uncl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vs.pack(side=tk.RIGHT, fill=tk.Y)

        log_top = tk.Frame(self.tab_log)
        log_top.pack(fill=tk.X, pady=2)
        _btn(log_top, "Обновить", self._load_log_into_ui, side=tk.LEFT, padx=2)
        _btn(log_top, "Открыть файл", self._open_log_file, side=tk.LEFT, padx=2)
        _btn(log_top, "Очистить журнал", self._clear_log_file, side=tk.LEFT, padx=2)
        log_wrap = tk.Frame(self.tab_log)
        log_wrap.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_wrap, height=16, wrap=tk.WORD)
        log_vs = ttk.Scrollbar(log_wrap, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_vs.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_vs.pack(side=tk.RIGHT, fill=tk.Y)

    def _make_tree(self, parent):
        wrap = tk.Frame(parent)
        wrap.pack(fill=tk.BOTH, expand=True)
        # Колонки задаём безопасными id (без точек) — иначе ttk на macOS рисует пусто.
        tree = ttk.Treeview(wrap, columns=("c0",), show="headings")
        tree.heading("c0", text="")
        tree.column("c0", width=120)
        vs = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=tree.yview)
        hs = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        tree.tag_configure("zero", foreground="#888888")
        tree.tag_configure("total", font=("Helvetica", 11, "bold"))
        wrap._tree = tree  # noqa: SLF001 — ссылка для пересоздания при необходимости
        return tree

    def _configure_tree_columns(self, tree, col_ids, headings, widths=None):
        """col_ids без точек/спецсимволов; headings — подписи."""
        tree.configure(columns=col_ids, displaycolumns=col_ids)
        for i, cid in enumerate(col_ids):
            tree.heading(cid, text=headings[i])
            w = 220 if i == 0 else 90
            if widths and i < len(widths):
                w = widths[i]
            tree.column(cid, width=w, anchor="w" if i == 0 else "center", stretch=True)

    def _set_date_widgets_state(self, state: str):
        for w in (self.start_entry, self.end_entry, self.start_cal_btn, self.end_cal_btn):
            try:
                w.config(state=state)
            except tk.TclError:
                pass

    def toggle_date_widgets(self):
        self._set_date_widgets_state("normal" if self.filter_enabled.get() else "disabled")
        if not self.store.ops.empty:
            self.run_analysis()

    def _parse_date(self, text: str):
        text = (text or "").strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        raise ValueError(f"Неверная дата: {text}")

    def _pick_date(self, variable: StringVar):
        """Календарь в отдельном окне — без DateEntry (он зависает на macOS Tk)."""
        if Calendar is None:
            messagebox.showinfo(
                "Календарь",
                "Пакет tkcalendar не установлен.\nВведите дату вручную: ДД.ММ.ГГГГ",
            )
            return
        try:
            current = self._parse_date(variable.get())
        except Exception:
            current = datetime.now().date()

        top = tk.Toplevel(self.root)
        top.title("Выбор даты")
        top.transient(self.root)
        top.resizable(False, False)
        # Тёмный фон: белые цифры, выходные жирным, выбранный день — голубым.
        cal = Calendar(
            top,
            selectmode="day",
            year=current.year,
            month=current.month,
            day=current.day,
            date_pattern="dd.mm.yyyy",
            background="#1e1e1e",
            foreground="white",
            bordercolor="#1e1e1e",
            headersbackground="#2a2a2a",
            headersforeground="white",
            normalbackground="#1e1e1e",
            normalforeground="white",
            weekendbackground="#1e1e1e",
            weekendforeground="white",
            othermonthbackground="#1e1e1e",
            othermonthforeground="#888888",
            othermonthwebackground="#1e1e1e",
            othermonthweforeground="#888888",
            selectbackground="#1a73e8",
            selectforeground="#7ec8ff",
            font=("Helvetica", 13),
        )
        try:
            p = cal._style_prefixe
            day_font = ("Helvetica", 13)
            we_font = ("Helvetica", 13, "bold")
            cal.style.configure(f"normal.{p}.TLabel", foreground="white", font=day_font)
            cal.style.configure(f"we.{p}.TLabel", foreground="white", font=we_font)
            cal.style.configure(f"normal_om.{p}.TLabel", foreground="#888888", font=day_font)
            cal.style.configure(f"we_om.{p}.TLabel", foreground="#888888", font=we_font)
            cal.style.configure(f"headers.{p}.TLabel", foreground="white")
            cal.style.configure(f"main.{p}.TLabel", foreground="white")
            # На aqua фон sel часто не виден — выделяем цветом текста.
            cal.style.configure(f"sel.{p}.TLabel", foreground="#7ec8ff", font=we_font)
        except Exception:
            pass
        cal.pack(padx=8, pady=8)
        # Повторно вешаем клик: на части macOS Tk событие <1> у ttk.Label теряется.
        for row in cal._calendar:
            for label in row:
                label.bind("<Button-1>", cal._on_click)

        applied = {"done": False}

        def apply_and_close(_event=None):
            if applied["done"]:
                return "break"
            applied["done"] = True
            try:
                d = cal.selection_get()
            except Exception:
                d = None
            if d is None:
                d = current
            variable.set(d.strftime("%d.%m.%Y"))
            try:
                top.destroy()
            except tk.TclError:
                pass
            if self.filter_enabled.get() and not self.store.ops.empty:
                self.run_analysis()
            return "break"

        # Клик по дню сразу выбирает дату (без обязательного OK).
        cal.bind("<<CalendarSelected>>", apply_and_close)

        bf = tk.Frame(top)
        bf.pack(pady=(0, 8))
        _btn(bf, "OK", apply_and_close, side=tk.LEFT, padx=4)
        _btn(bf, "Отмена", top.destroy, side=tk.LEFT, padx=4)
        top.protocol("WM_DELETE_WINDOW", top.destroy)
        # Без grab_set — на macOS он часто блокирует клики по дням.
        top.focus_force()
        top.lift()

    def _on_year_changed(self):
        try:
            year = int(self.year_var.get())
        except ValueError:
            return
        self.summary_cfg["year"] = year
        suggested = suggest_summary_path(APP_DIR, year)
        if suggested.exists():
            self.summary_path.set(str(suggested))
            self.log_message(f"Год {year}: сводная {suggested.name}")
        else:
            self.log_message(
                f"Год {year}: файла {suggested.name} нет — создайте через «Создать сводную на год…»",
                level="WARNING",
            )
        if not self.store.ops.empty:
            self.run_analysis()

    def _tree_to_tsv(self, tree) -> str:
        cols = list(tree["columns"] or ())
        if not cols:
            return ""
        headers = [tree.heading(c)["text"] for c in cols]
        lines = ["\t".join(headers)]
        for item in tree.get_children():
            vals = tree.item(item).get("values") or []
            lines.append("\t".join("" if v is None else str(v) for v in vals))
        return "\n".join(lines)

    def copy_preview(self):
        parts = []
        t_cat = self._tree_to_tsv(getattr(self, "tree_preview_cat", None))
        t_tot = self._tree_to_tsv(getattr(self, "tree_preview_tot", None))
        t_form = self._tree_to_tsv(getattr(self, "tree_form", None))
        if t_cat:
            parts.append("Категории по неделям\n" + t_cat)
        if t_tot:
            parts.append("Итоги по неделям\n" + t_tot)
        if t_form:
            parts.append("Форма 4001\n" + t_form)
        text = "\n\n".join(parts)
        if not text.strip():
            messagebox.showinfo("Копирование", "Нечего копировать — сначала обновите превью")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()
        self._preview_clipboard = text
        self.status_var.set("Превью скопировано в буфер")
        self.log_message("Превью скопировано в буфер обмена")

    def _copy_focused_tree(self, event=None):
        # Если активна вкладка превью — копируем всё превью одной пачкой.
        try:
            tab_text = self.notebook.tab(self.notebook.select(), "text")  # type: ignore[arg-type]
        except Exception:
            tab_text = ""
        if str(tab_text).startswith("Превью"):
            self.copy_preview()
            return "break"

        w = self.root.focus_get()
        for t in (
            getattr(self, "tree_preview_cat", None),
            getattr(self, "tree_preview_tot", None),
            getattr(self, "tree_form", None),
            self.tree_uncl,
        ):
            if t is None:
                continue
            if str(w).startswith(str(t)):
                text = self._tree_to_tsv(t)
                if text:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(text)
                    self.status_var.set("Таблица скопирована")
                return "break"
        return None

    def _load_log_into_ui(self):
        """При открытии — показать хвост analysis.log и прокрутить к последнему событию."""
        self.log_text.delete("1.0", tk.END)
        lines = APP_LOG.read_lines()
        if lines:
            self.log_text.insert(tk.END, "\n".join(lines) + "\n")
            self.log_text.insert(tk.END, "—" * 40 + "\n")
        self.log_text.see(tk.END)
        self.root.after(100, lambda: self.log_text.see(tk.END))

    def log_message(self, msg: str, tag: str = None, level: str = "INFO"):
        line = APP_LOG.append(msg, level=level)
        if hasattr(self, "log_text"):
            self.log_text.insert(tk.END, line + "\n")
            self.log_text.see(tk.END)
        logging.log(getattr(logging, level, logging.INFO), msg)

    def _on_tab_changed(self, _event=None):
        try:
            if self.notebook.index(self.notebook.select()) == self.notebook.index(self.tab_log):
                self.log_text.see(tk.END)
        except Exception:
            pass

    def _busy(self, on: bool):
        self.root.config(cursor="watch" if on else "")
        self.root.update_idletasks()

    def _set_kpis_empty(self):
        for v in (
            self.kpi_ops_var,
            self.kpi_patients_var,
            self.kpi_plan_var,
            self.kpi_emerg_var,
            self.kpi_period_var,
            self.kpi_files_var,
            self.kpi_diff_var,
        ):
            v.set("—")

    def _refresh_sources_list(self):
        self.sources_text.configure(state=tk.NORMAL)
        self.sources_text.delete("1.0", tk.END)
        if self.store.ops.empty:
            self.sources_text.insert(tk.END, "Нет загруженных журналов")
        else:
            self.store.refresh_source_meta()
            blocks = []
            for src in self.store.sources:
                meta = self.store.source_meta.get(src) or {}
                d0 = meta.get("date_from")
                d1 = meta.get("date_to")
                n = meta.get("count", 0)
                if d0 is not None and d1 is not None:
                    period = f"{d0.strftime('%d.%m.%Y')} – {d1.strftime('%d.%m.%Y')}"
                    blocks.append(f"• {src}\n  {period}\n  операций: {n}")
                else:
                    blocks.append(f"• {src}\n  операций: {n}")
            self.sources_text.insert(tk.END, "\n\n".join(blocks))
        self.sources_text.configure(state=tk.DISABLED)
        if self.df_emk is not None:
            name = os.path.basename(self.emk_path) if self.emk_path else "загружен"
            self.emk_status.config(text=f"ЭМК: {name} ({len(self.df_emk)} стр.)")
        else:
            self.emk_status.config(text="ЭМК: не загружен")
        self.kpi_files_var.set(f"{len(self.store.sources)} / {'да' if self.df_emk is not None else 'нет'}")

    def _on_source_dblclick(self, _e=None):
        # строка под курсором
        try:
            idx = self.sources_text.index(f"@{_e.x},{_e.y}") if _e else self.sources_text.index(tk.INSERT)
            line = self.sources_text.get(f"{idx} linestart", f"{idx} lineend").strip()
        except Exception:
            line = ""
        if line:
            self.log_message(f"Источник: {line}")
            self.notebook.select(self.tab_log)

    def _on_plan_mode_change(self):
        if not self.store.ops.empty:
            self.run_analysis()

    def _weeks_for_month(self, year: int, month: int):
        path = self.summary_path.get().strip()
        sheet_names = self.summary_cfg.get("sheet_names", {})
        sheet = sheet_names.get(month) or sheet_names.get(str(month))
        if path and os.path.exists(path) and sheet:
            try:
                wb = openpyxl.load_workbook(path, data_only=False)
                if sheet in wb.sheetnames:
                    weeks = read_sheet_weeks(wb[sheet])
                    wb.close()
                    if weeks:
                        return weeks
                wb.close()
            except Exception:
                pass
        return compute_month_weeks(year, month)

    def choose_summary(self):
        path = filedialog.askopenfilename(initialdir=str(APP_DIR), filetypes=[("Excel", "*.xlsx")])
        if path:
            self.summary_path.set(path)
            self.log_message(f"Сводная: {path}")

    def clear_store(self):
        if messagebox.askyesno("Очистка", "Удалить накопленные операции?"):
            self.store.clear()
            self.cat_table = self.totals_df = None
            self.weeks = []
            self._refresh_sources_list()
            self._set_kpis_empty()
            self._clear_trees()
            self.log_message("Накопитель очищен")
            self.status_var.set("Очищено")

    def _clear_trees(self):
        for tree in (
            getattr(self, "tree_preview_cat", None),
            getattr(self, "tree_preview_tot", None),
            getattr(self, "tree_form", None),
            self.tree_uncl,
        ):
            if tree is None:
                continue
            for item in tree.get_children():
                tree.delete(item)

    def load_emk(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not path:
            return
        try:
            self._busy(True)
            self.df_emk = read_table(path)
            self.emk_path = path
            self._refresh_sources_list()
            self.log_message(f"ЭМК: {path} ({len(self.df_emk)} строк)")
            if not self.store.ops.empty:
                n = self._rebind_emk_to_store()
                self.log_message(f"Привязка ЭМК: {n} операций")
                self.run_analysis()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            logging.error(traceback.format_exc())
        finally:
            self._busy(False)

    def _rebind_emk_to_store(self) -> int:
        if self.df_emk is None or self.store.ops.empty:
            return 0
        from analyzers.surgery import SurgeryAnalyzer as SA

        analyzer = SA.__new__(SA)
        analyzer.emk_df = self.df_emk
        analyzer.categories = self.config["surgery_categories"]
        analyzer._code_index = analyzer._build_code_index()
        analyzer._emk_hosp_index = analyzer._build_emk_hosp_index()
        updated = 0
        ops = self.store.ops
        for i, row in ops.iterrows():
            hosp = analyzer.hosp_type_for(row["КВС"], row["Дата"])
            diag = analyzer.diagnosis_for(row["КВС"], row["Дата"])
            changed = False
            if (hosp or "") != str(row.get("Тип_ЭМК") or ""):
                ops.at[i, "Тип_ЭМК"] = hosp or ""
                changed = True
            if (diag or "") != str(row.get("Диагноз") or ""):
                ops.at[i, "Диагноз"] = diag or ""
                changed = True
            cat = str(row.get("Категория") or "")
            if cat in ("Миринготомия план", "Миринготомия экстр"):
                day = pd.Timestamp(row["Дата"]).normalize()
                has_adeno = (
                    (ops["КВС"].astype(str) == str(row["КВС"]))
                    & (pd.to_datetime(ops["Дата"]).dt.normalize() == day)
                    & (ops["Категория"] == "Аденотомия")
                ).any()
                new_cat = (
                    "Миринготомия план"
                    if has_adeno
                    else ("Миринготомия экстр" if hosp == "экстренная" else "Миринготомия план")
                )
                if new_cat != cat:
                    ops.at[i, "Категория"] = new_cat
                    changed = True
            if changed:
                updated += 1
        return updated

    def load_surg(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not paths:
            return
        try:
            self._busy(True)
            for path in paths:
                try:
                    df = read_table(path)
                    analyzer = SurgeryAnalyzer(
                        df, self.dept_var.get(), self.config["surgery_categories"], emk_df=self.df_emk
                    )
                    ops = analyzer.extract_operations()
                    if ops.empty:
                        self.log_message(f"Нет операций: {os.path.basename(path)}")
                        continue
                    info = self.store.add(ops, path)
                    self.last_batch_span = (info.get("date_from"), info.get("date_to"))
                    msg = f"{os.path.basename(path)}: +{info['added']}, вытеснено {info['removed']}, всего {info['total']}"
                    if info.get("date_from") is not None:
                        msg += f" | {info['date_from'].strftime('%d.%m.%Y')}–{info['date_to'].strftime('%d.%m.%Y')}"
                    self.log_message(msg)
                    uncl = ops[ops["Категория"] == "Не классифицировано"]
                    if len(uncl):
                        codes = sorted({c for c in uncl["Код"].dropna().unique() if c})
                        self.log_message(f"  не классифицировано: {len(uncl)} опер., коды: {codes}")
                        for code in codes:
                            hint = ""
                            if "КСГ_подсказка" in uncl.columns:
                                rows = uncl[uncl["Код"] == code]
                                if len(rows):
                                    hint = str(rows.iloc[0].get("КСГ_подсказка", "") or "")
                            if hint:
                                self.log_message(f"    {code}: {hint}")
                            else:
                                self.log_message(f"    {code}: нет в KSGoperacii.csv — добавьте в config.yaml")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"{os.path.basename(path)}:\n{e}")
                    logging.error(traceback.format_exc())
            self._refresh_sources_list()
            self.run_analysis()
            self.notebook.select(self.tab_preview)
        finally:
            self._busy(False)

    def get_view_ops(self) -> pd.DataFrame:
        ops = self.store.ops.copy()
        if ops.empty:
            return ops
        if self.filter_enabled.get():
            try:
                start = pd.Timestamp(self._parse_date(self.start_date_var.get()))
                end = pd.Timestamp(self._parse_date(self.end_date_var.get())) + timedelta(days=1) - timedelta(seconds=1)
            except Exception:
                messagebox.showerror("Дата", "Формат даты: ДД.ММ.ГГГГ")
                return ops
            dates = pd.to_datetime(ops["Дата"], errors="coerce")
            ops = ops.loc[(dates >= start) & (dates <= end)].copy()
        return ops

    def run_analysis(self):
        ops = self.get_view_ops()
        if ops.empty:
            messagebox.showwarning("Нет данных", "Добавьте опержурнал(ы)")
            return
        try:
            self._busy(True)
            cat_table, totals_df, weeks = build_summary_tables(
                ops, self.summary_cfg, self.config["surgery_categories"]
            )
            if self.plan_mode.get() == "emk" and self.df_emk is not None:
                totals_df = self._totals_from_emk(ops, totals_df, weeks)
            self.cat_table, self.totals_df, self.weeks = cat_table, totals_df, weeks
            self._update_month_choices(ops)
            self.refresh_preview()
            self._update_unclassified(ops)
            self._update_kpis(ops, totals_df)
            self._maybe_update_emk_kpi(ops)
            self.status_var.set(f"Готово: {len(ops)} операций, {len(weeks)} нед.")
            d0 = pd.to_datetime(ops["Дата"]).min()
            d1 = pd.to_datetime(ops["Дата"]).max()
            prev_cat_rows = len(self.tree_preview_cat.get_children()) if hasattr(self, "tree_preview_cat") else 0
            prev_tot_rows = len(self.tree_preview_tot.get_children()) if hasattr(self, "tree_preview_tot") else 0
            self.log_message(
                f"Отчёт: {len(ops)} опер. | {d0.strftime('%d.%m.%Y')}–{d1.strftime('%d.%m.%Y')} | "
                f"месяц превью={self.preview_month.get()!r} | категории превью={prev_cat_rows} | итоги превью={prev_tot_rows}"
            )
        except Exception as e:
            self.log_message(traceback.format_exc(), level="ERROR")
            messagebox.showerror("Ошибка", str(e))
        finally:
            self._busy(False)

    def _update_month_choices(self, ops):
        months = sorted({int(m) for m in pd.to_datetime(ops["Дата"]).dt.month.dropna().unique()})
        try:
            year = int(self.year_var.get())
        except ValueError:
            year = int(self.summary_cfg.get("year", 2026))
        self.summary_cfg["year"] = year
        labels = []
        self._month_label_to_num = {}
        for m in months:
            if m not in MONTH_RU:
                continue
            lab = f"{MONTH_RU[m].capitalize()} {year}"
            labels.append(lab)
            self._month_label_to_num[lab] = m
        self.month_combo["values"] = labels
        if labels:
            prefer = None
            if self.last_batch_span[0] is not None:
                prefer = f"{MONTH_RU[int(self.last_batch_span[0].month)].capitalize()} {year}"
            self.preview_month.set(prefer if prefer in labels else labels[0])
        else:
            self.preview_month.set("")
            self.log_message("Превью: не удалось определить месяц по датам операций", level="WARNING")

    def _totals_from_emk(self, ops, totals_df, weeks):
        df = totals_df.copy()
        emerg = pd.Series(0, index=weeks, dtype=int)
        plan = pd.Series(0, index=weeks, dtype=int)
        ops = ops.copy()
        ops["Неделя_начало"] = pd.to_datetime(ops["Дата"]).apply(lambda d: d - pd.Timedelta(days=d.weekday()))
        for _, row in ops.iterrows():
            w = row["Неделя_начало"]
            t = str(row.get("Тип_ЭМК", "") or "").lower()
            if t.startswith("экстр"):
                emerg[w] = emerg.get(w, 0) + 1
            elif t.startswith("план"):
                plan[w] = plan.get(w, 0) + 1
        df.loc["Экстренно операций"] = emerg.reindex(weeks, fill_value=0).astype(int).values
        df.loc["План операций"] = plan.reindex(weeks, fill_value=0).astype(int).values
        return df.astype(int)

    def _update_kpis(self, ops, totals_df):
        self.kpi_ops_var.set(str(len(ops)))
        self.kpi_patients_var.set(str(int(ops["КВС"].nunique())) if len(ops) else "0")
        if totals_df is not None and not totals_df.empty and "Всего операций" in totals_df.index:
            total = int(totals_df.loc["Всего операций"].sum())
            plan = int(totals_df.loc["План операций"].sum()) if "План операций" in totals_df.index else 0
            emerg = int(totals_df.loc["Экстренно операций"].sum()) if "Экстренно операций" in totals_df.index else 0
            self.kpi_plan_var.set(f"{plan * 100 / total:.0f}%" if total else "—")
            self.kpi_emerg_var.set(f"{emerg * 100 / total:.0f}%" if total else "—")
        if len(ops):
            d0 = pd.to_datetime(ops["Дата"]).min().strftime("%d.%m.%Y")
            d1 = pd.to_datetime(ops["Дата"]).max().strftime("%d.%m.%Y")
            self.kpi_period_var.set(f"{d0} – {d1}")

    def _maybe_update_emk_kpi(self, ops):
        if self.df_emk is None or ops.empty:
            self.kpi_diff_var.set("—")
            return
        if ops["Тип_ЭМК"].astype(str).str.strip().eq("").all():
            self.kpi_diff_var.set("нет связи")
            return
        result = compare_plan_emergency(ops, self.summary_cfg)
        self.last_emk_compare = result
        self.kpi_diff_var.set(str(len(result["mismatches"])))

    def on_emk_mode(self):
        if self.df_emk is None:
            messagebox.showinfo("ЭМК", "Сначала загрузите ЭМК")
            self.plan_mode.set("template")
            return
        self.run_analysis()
        self.show_emk_diff()

    def show_emk_diff(self):
        ops = self.get_view_ops()
        if ops.empty:
            messagebox.showwarning("Нет данных", "Нет операций")
            return
        if self.df_emk is None:
            messagebox.showwarning("Нет ЭМК", "Загрузите ЭМК")
            return
        if ops["Тип_ЭМК"].astype(str).str.strip().eq("").all():
            messagebox.showwarning("Нет связи", "КВС журнала и ЭМК не пересекаются")
            return
        result = compare_plan_emergency(ops, self.summary_cfg)
        self.kpi_diff_var.set(str(len(result["mismatches"])))
        self.notebook.select(self.tab_log)
        self.log_message(format_mismatch_report(result))
        if result["mismatches"]:
            messagebox.showwarning("Расхождения", f"Найдено: {len(result['mismatches'])}\nСм. вкладку Журнал")
        else:
            messagebox.showinfo("Сверка", "Расхождений нет")

    def display_tables(self):
        """Раньше заполнял дублирующие вкладки Категории/Итоги — оставлены только превью."""
        return

    def _fill_tree(self, tree, first_col, week_headers, df, week_keys, totals=False):
        for item in tree.get_children():
            tree.delete(item)
        if df is None or df.empty:
            return
        col_ids = ["c0"] + [f"w{i}" for i in range(len(week_headers))]
        headings = [first_col] + list(week_headers)
        self._configure_tree_columns(tree, col_ids, headings)
        for idx, row in df.iterrows():
            values = [idx] + [int(row[week]) for week in week_keys]
            tags = []
            if totals or idx in (
                "Всего операций", "Экстренно операций", "План операций", "Дети всего", "Взрослые", "Человек"
            ):
                tags.append("total")
            elif sum(values[1:]) == 0:
                tags.append("zero")
            tree.insert("", tk.END, values=values, tags=tuple(tags))

    def refresh_preview(self):
        # очищаем 3 таблицы превью
        for tree in (self.tree_preview_cat, self.tree_preview_tot, self.tree_form):
            for item in tree.get_children():
                tree.delete(item)
        ops = self.get_view_ops()
        if ops.empty:
            self.preview_info.set("Превью: нет операций")
            return
        label = self.preview_month.get()
        month_map = getattr(self, "_month_label_to_num", {})
        if not label or label not in month_map:
            self.preview_info.set(f"Превью: месяц не выбран (label={label!r})")
            self.log_message(f"Превью пусто: label={label!r}, map={list(month_map.keys())}", level="WARNING")
            return
        month = int(month_map[label])
        try:
            year = int(self.year_var.get())
        except ValueError:
            year = int(self.summary_cfg.get("year", 2026))
        weeks = self._weeks_for_month(year, month)
        if not weeks:
            weeks = compute_month_weeks(year, month)
        week_headers = [f"{s.strftime('%d.%m')}-{e.strftime('%d.%m')}" for s, e in weeks]
        # --- превью категории ---
        col_ids_cat = ["c0"] + [f"w{i}" for i in range(len(weeks))] + ["tot"]
        headings_cat = ["Категория"] + week_headers + ["ИТОГ"]
        self._configure_tree_columns(
            self.tree_preview_cat, col_ids_cat, headings_cat, widths=[240] + [90] * (len(weeks) + 1)
        )

        month_ops = ops[pd.to_datetime(ops["Дата"]).dt.month == month].copy()
        cat_order = list(self.summary_cfg.get("category_rows", {}).keys())
        counts_map = {cat: [0] * len(weeks) for cat in cat_order}
        unmapped = 0
        for _, r in month_ops.iterrows():
            cat = r["Категория"]
            if cat not in counts_map:
                continue
            d = pd.Timestamp(r["Дата"]).date()
            hit = False
            for wi, (s, e) in enumerate(weeks):
                if s <= d <= e:
                    counts_map[cat][wi] += 1
                    hit = True
                    break
            if not hit:
                unmapped += 1

        rows_cat = 0
        for cat in cat_order:
            counts = counts_map[cat]
            total = sum(counts)
            if self.hide_zeros.get() and total == 0:
                continue
            self.tree_preview_cat.insert(
                "", tk.END, values=[cat] + counts + [total], tags=("zero",) if total == 0 else ()
            )
            rows_cat += 1

        emerg_set = set(self.summary_cfg.get("emergency_categories", []))
        plan_set = set(self.summary_cfg.get("plan_categories", []))

        def week_slice(wi):
            s, e = weeks[wi]
            return month_ops[month_ops["Дата"].map(lambda x: s <= pd.Timestamp(x).date() <= e)]

        arrays = {k: [] for k in ("ops", "emerg", "plan", "kids", "people")}
        for i in range(len(weeks)):
            wops = week_slice(i)
            arrays["ops"].append(len(wops))
            arrays["emerg"].append(int(wops["Категория"].isin(emerg_set).sum()))
            arrays["plan"].append(int(wops["Категория"].isin(plan_set).sum()))
            arrays["kids"].append(int(wops.loc[wops["Возраст"].fillna(99) < 18, "КВС"].nunique()))
            arrays["people"].append(int(wops["КВС"].nunique()))
        adults = [max(0, arrays["people"][i] - arrays["kids"][i]) for i in range(len(weeks))]

        # --- превью итоги ---
        col_ids_tot = ["c0"] + [f"w{i}" for i in range(len(weeks))] + ["tot"]
        headings_tot = ["Показатель"] + week_headers + ["ИТОГ"]
        self._configure_tree_columns(
            self.tree_preview_tot, col_ids_tot, headings_tot, widths=[260] + [90] * (len(weeks) + 1)
        )
        rows_tot = 0
        for name, arr in (
            ("Всего операций", arrays["ops"]),
            ("Экстренно операций", arrays["emerg"]),
            ("План операций", arrays["plan"]),
            ("Дети всего", arrays["kids"]),
            ("Взрослые", adults),
            ("Человек", arrays["people"]),
        ):
            self.tree_preview_tot.insert("", tk.END, values=[name] + arr + [sum(arr)], tags=("total",))
            rows_tot += 1

        # --- форма 4001 ---
        form_cfg = self.summary_cfg.get("form_4001") or {}
        stats = compute_form_4001(
            month_ops,
            self.config.get("surgery_categories", []),
            pension_age=int(self.config.get("thresholds", {}).get("pension_age", 60)),
        )
        # Колонки как в шаблоне: 1,2,3,4,5,6,28 + S (т.4000)
        form_cols = ["c0", "line", "n", "o", "p", "q", "r", "s"]
        form_heads = [
            "Наименование операции",
            "№ строки",
            "всего",
            "0–14 лет",
            "до 1 года",
            "15–17 лет",
            "морфол. (28)",
            "Всего (т.4000)",
        ]
        self._configure_tree_columns(
            self.tree_form, form_cols, form_heads, widths=[280, 70, 70, 70, 70, 70, 90, 90]
        )
        for row in form_4001_preview_rows(stats):
            name = row["name"]
            bold = name in (
                "операции на органах уха, горла, носа",
                "операции на органах дыхания",
                "операции на коже и подкожной клетчатке",
                "Всего операций",
            )
            self.tree_form.insert(
                "",
                tk.END,
                values=(
                    name,
                    row.get("line", ""),
                    row.get("total", ""),
                    row.get("age_0_14", ""),
                    row.get("age_under_1", ""),
                    row.get("age_15_17", ""),
                    row.get("histology", ""),
                    row.get("senior", ""),
                ),
                tags=("total",) if bold else (),
            )

        self.preview_info.set(
            f"Превью: {label} | операций месяца {len(month_ops)} | недель {len(weeks)} | категории {rows_cat} | итоги {rows_tot}"
            + (f" | вне недель: {unmapped}" if unmapped else "")
        )
        self.log_message(
            f"Превью: {label}, недель={[(str(s), str(e)) for s, e in weeks]}, "
            f"ops_month={len(month_ops)}, categories={rows_cat}, totals={rows_tot}, unmapped={unmapped}, "
            f"form4001={stats}"
        )

    def _update_unclassified(self, ops):
        for item in self.tree_uncl.get_children():
            self.tree_uncl.delete(item)
        uncl = ops[ops["Категория"] == "Не классифицировано"]
        self.notebook.tab(self.tab_uncl, text=f"Не классифицировано ({len(uncl)})")
        for _, r in uncl.iterrows():
            dt = r["Дата"]
            dt_s = dt.strftime("%d.%m.%Y") if hasattr(dt, "strftime") else str(dt)
            svc = str(r.get("Услуга", "") or "")[:80]
            self.tree_uncl.insert(
                "",
                tk.END,
                values=(
                    dt_s,
                    r.get("КВС"),
                    r.get("Код"),
                    r.get("КСГ_название", ""),
                    r.get("КСГ", ""),
                    svc,
                ),
            )

    def update_summary(self):
        """Один диалог записи: галочки «Недели» и «Форма 4001»."""
        if self.store.ops.empty:
            messagebox.showwarning("Нет данных", "Сначала добавьте опержурнал(ы)")
            return
        path = self.summary_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Нет файла", f"Сводная не найдена:\n{path}")
            return

        d_min, d_max = self.last_batch_span
        if d_min is None or d_max is None:
            d_min, d_max = self.store.date_span(self.store.ops)
        try:
            cfg_year = int(self.year_var.get())
        except ValueError:
            cfg_year = int(self.summary_cfg.get("year", 2026))

        if d_min is not None:
            op_years = sorted({int(y) for y in pd.to_datetime(self.store.ops["Дата"]).dt.year.dropna().unique()})
            foreign = [y for y in op_years if y != cfg_year]
            if foreign:
                if not messagebox.askyesno(
                    "Год данных",
                    f"В накопителе есть операции за {op_years}, а сводная/год в настройках — {cfg_year}.\n\n"
                    f"Записать всё равно в:\n{path}\n\n"
                    "Для другого года: смените «Год» или «Создать сводную на год…».",
                ):
                    return

        period = ""
        if d_min is not None:
            period = f"{d_min.strftime('%d.%m.%Y')} – {d_max.strftime('%d.%m.%Y')}"

        top = tk.Toplevel(self.root)
        top.title("Запись в Excel")
        top.transient(self.root)
        top.resizable(False, False)
        tk.Label(top, text=f"Файл:\n{path}", justify=tk.LEFT, wraplength=480).pack(padx=12, pady=(12, 4), anchor="w")
        if period:
            tk.Label(top, text=f"Период перезаписи: {period}", justify=tk.LEFT).pack(padx=12, pady=2, anchor="w")
        tk.Label(top, text=f"Год настроек: {cfg_year}", justify=tk.LEFT).pack(padx=12, pady=2, anchor="w")

        weeks_var = BooleanVar(value=self.write_weeks_var.get())
        form_var = BooleanVar(value=self.write_form_var.get())
        tk.Checkbutton(top, text="Недели / категории (столбцы C–G)", variable=weeks_var).pack(
            padx=12, pady=4, anchor="w"
        )
        tk.Checkbutton(top, text="Форма 4001 (формулы N/R не затираются)", variable=form_var).pack(
            padx=12, pady=4, anchor="w"
        )
        tk.Label(
            top,
            text="Закройте файл в Excel перед записью. Создаётся .bak с датой.",
            fg="#555",
            wraplength=480,
            justify=tk.LEFT,
        ).pack(padx=12, pady=6, anchor="w")

        def do_write():
            write_weeks = bool(weeks_var.get())
            write_form = bool(form_var.get())
            if not write_weeks and not write_form:
                messagebox.showwarning("Запись", "Отметьте хотя бы один пункт", parent=top)
                return
            self.write_weeks_var.set(write_weeks)
            self.write_form_var.set(write_form)
            top.destroy()
            self._do_write_excel(path, d_min, d_max, write_weeks=write_weeks, write_form=write_form)

        bf = tk.Frame(top)
        bf.pack(pady=10)
        _btn(bf, "Записать", do_write, side=tk.LEFT, padx=4)
        _btn(bf, "Отмена", top.destroy, side=tk.LEFT, padx=4)
        top.grab_set()
        top.focus_force()

    def _do_write_excel(self, path, d_min, d_max, *, write_weeks: bool, write_form: bool):
        if excel_file_locked(path):
            messagebox.showerror(
                "Файл занят",
                f"Сводная открыта в Excel или заблокирована:\n{path}\n\n"
                "Закройте файл и повторите запись.",
            )
            return
        try:
            self._busy(True)
            parts = []
            if write_weeks:
                parts.append("недели")
            if write_form:
                parts.append("форма 4001")
            self.log_message(
                f"Запись в сводную ({', '.join(parts)}): {path} | ops={len(self.store.ops)} | "
                f"overwrite={d_min}..{d_max}"
            )
            writer = SummaryWriter(
                path,
                self.summary_cfg,
                department=self.dept_var.get(),
                categories=self.config.get("surgery_categories", []),
                pension_age=int(self.config.get("thresholds", {}).get("pension_age", 60)),
            )
            report = writer.write(
                self.store.ops,
                output_path=path,
                overwrite_from=d_min,
                overwrite_to=d_max,
                backup=True,
                write_weeks=write_weeks,
                write_form=write_form,
            )
            months = ", ".join(report.get("months", {}).keys()) or "—"
            self.log_message(
                f"Сводная обновлена: {months}, ячеек {report.get('cells_written', 0)}, "
                f"вне недель {report.get('unmapped_dates', 0)}"
            )
            for sheet, info in report.get("months", {}).items():
                self.log_message(
                    f"  [{sheet}] cols={info.get('cols')} weeks={info.get('weeks')} "
                    f"ops={info.get('ops')} sample={info.get('sample')}"
                )
                form = info.get("form_4001")
                if form:
                    self.log_message(f"  [{sheet}] форма 4001: {form}")
            self.status_var.set(f"Запись: {report.get('cells_written', 0)} ячеек")
            self._persist_settings()
            if messagebox.askyesno(
                "Готово",
                f"Обновлено: {months}\nЯчеек: {report.get('cells_written', 0)}\n\n"
                "Открыть файл? (если Excel был открыт — закройте и откройте заново)",
            ):
                self.open_summary_file()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.log_message(traceback.format_exc(), level="ERROR")
        finally:
            self._busy(False)

    def open_summary_file(self):
        path = self.summary_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Нет файла", path)
            return
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def export_simple(self):
        if self.store.ops.empty:
            messagebox.showwarning("Нет данных", "Сначала анализ / загрузка журнала")
            return
        path_tpl = self.summary_path.get().strip()
        if not path_tpl or not os.path.exists(path_tpl):
            messagebox.showerror("Нет шаблона", f"Нужен файл сводной как образец стиля:\n{path_tpl}")
            return
        label = self.preview_month.get()
        month_map = getattr(self, "_month_label_to_num", {})
        month = int(month_map[label]) if label in month_map else None
        try:
            year = int(self.year_var.get())
        except ValueError:
            year = int(self.summary_cfg.get("year", 2026))
        default_name = f"Отчёт_{MONTH_RU.get(month or 0, 'период')}_{year}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_path:
            return
        try:
            self._busy(True)
            report = export_month_like_summary(
                path_tpl,
                file_path,
                self.get_view_ops(),
                self.summary_cfg,
                department=self.dept_var.get(),
                categories=self.config.get("surgery_categories", []),
                pension_age=int(self.config.get("thresholds", {}).get("pension_age", 60)),
                month=month,
                year=year,
            )
            self.log_message(
                f"Экспорт как сводная: {file_path} | ячеек {report.get('cells_written', 0)} | "
                f"листы {list(report.get('months', {}))}"
            )
            messagebox.showinfo(
                "Готово",
                f"Сохранён отчёт в стиле сводной:\n{file_path}\n\n"
                "Листы/формулы как в шаблоне; заполнен выбранный месяц.",
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.log_message(traceback.format_exc(), level="ERROR")
        finally:
            self._busy(False)

    def create_year_summary_dialog(self):
        """Создать файл «Операции сводная YYYY.xlsx» на новый год из текущего шаблона."""
        tpl = self.summary_path.get().strip()
        if not tpl or not os.path.exists(tpl):
            messagebox.showerror("Нет шаблона", "Укажите существующую сводную как основу")
            return
        try:
            cur = int(self.year_var.get())
        except ValueError:
            cur = 2026
        new_year = cur + 1
        # простой диалог
        top = tk.Toplevel(self.root)
        top.title("Сводная на год")
        top.transient(self.root)
        tk.Label(top, text="Создать файл сводной на год:").pack(padx=12, pady=8)
        yvar = StringVar(value=str(new_year))
        ttk.Combobox(top, textvariable=yvar, values=[str(y) for y in range(cur, cur + 6)], width=8).pack()
        out_hint = StringVar(value=str(suggest_summary_path(APP_DIR, new_year)))

        def refresh_hint(*_a):
            try:
                out_hint.set(str(suggest_summary_path(APP_DIR, int(yvar.get()))))
            except ValueError:
                pass

        yvar.trace_add("write", refresh_hint)
        tk.Label(top, textvariable=out_hint, wraplength=420, justify=tk.LEFT).pack(padx=12, pady=4)

        def do_create():
            try:
                y = int(yvar.get())
            except ValueError:
                messagebox.showerror("Год", "Укажите год числом", parent=top)
                return
            dest = Path(out_hint.get())
            if dest.exists() and not messagebox.askyesno(
                "Файл есть", f"{dest.name} уже существует. Перезаписать?", parent=top
            ):
                return
            try:
                path = create_year_summary(
                    tpl,
                    y,
                    output_path=str(dest),
                    sheet_names=self.summary_cfg.get("sheet_names"),
                    clear_values=True,
                )
                self.year_var.set(str(y))
                self.summary_cfg["year"] = y
                self.summary_path.set(str(path))
                self.log_message(f"Создана сводная на {y}: {path}")
                messagebox.showinfo("Готово", f"Создан файл:\n{path}", parent=top)
                top.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=top)

        bf = tk.Frame(top)
        bf.pack(pady=10)
        _btn(bf, "Создать", do_create, side=tk.LEFT, padx=4)
        _btn(bf, "Отмена", top.destroy, side=tk.LEFT, padx=4)

    def show_about(self):
        ver = read_local_version(APP_DIR)
        upd = self.config.get("updates") or {}
        repo = f"{upd.get('github_owner', '')}/{upd.get('github_repo', '')}".strip("/")
        repo_line = f"\nGitHub: {repo}" if repo != "/" and repo else ""
        messagebox.showinfo(
            "О программе",
            f"Сводная операций v{ver}\nВидновская КБ{repo_line}\n\n"
            "Запись: один диалог с галочками «Недели» и «Форма 4001».\n"
            "Обновления: Помощь → Проверить обновления…",
        )

    def check_updates(self, silent: bool = False):
        """Проверить GitHub и предложить установить обновление."""
        cfg = self.config.get("updates") or {}
        if not cfg.get("enabled", True):
            if not silent:
                messagebox.showinfo("Обновления", "Проверка обновлений отключена в config.yaml")
            return
        try:
            self._busy(True)
            self.status_var.set("Проверка обновлений…")
            info = check_for_update(APP_DIR, cfg)
        except Exception as e:
            self.log_message(f"Обновление: ошибка проверки — {e}", level="WARNING")
            if not silent:
                messagebox.showerror("Обновления", str(e))
            return
        finally:
            self._busy(False)
            self.status_var.set("Готов к работе")

        if info is None:
            ver = read_local_version(APP_DIR)
            msg = f"У вас актуальная версия: {ver}"
            self.log_message(f"Обновления: {msg}")
            if not silent:
                messagebox.showinfo("Обновления", msg)
            return

        self.log_message(
            f"Обновления: доступна {info.remote_version} (сейчас {info.local_version})"
        )
        notes = format_update_notes(info)
        top = tk.Toplevel(self.root)
        top.title("Доступно обновление")
        top.transient(self.root)
        top.resizable(True, True)
        tk.Label(top, text="Найдена новая версия", font=("Helvetica", 13, "bold")).pack(
            padx=12, pady=(12, 4), anchor="w"
        )
        txt = tk.Text(top, width=64, height=14, wrap=tk.WORD)
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        txt.insert("1.0", notes)
        txt.configure(state=tk.DISABLED)
        include_cfg = BooleanVar(value=False)
        tk.Checkbutton(
            top,
            text="Также заменить config.yaml (рубрикатор) — обычно не нужно",
            variable=include_cfg,
        ).pack(padx=12, pady=4, anchor="w")

        def do_install():
            if not messagebox.askyesno(
                "Обновление",
                "Установить обновление?\n\n"
                "Будут обновлены код и VERSION.\n"
                "Excel, журналы, .venv и ui_settings.json не трогаются.\n"
                "Создаётся папка .update_backup_…",
                parent=top,
            ):
                return
            try:
                self._busy(True)
                self.status_var.set("Загрузка обновления…")
                report = apply_update_from_zip(
                    APP_DIR,
                    info.zip_url,
                    token=resolve_token(cfg),
                    include_config=bool(include_cfg.get()),
                    backup=True,
                    sha256_url=info.sha256_url,
                    require_sha256=bool(info.source == "release-asset"),
                    mode="release-asset" if info.source == "release-asset" else "auto",
                )
                self.log_message(
                    f"Обновление установлено: v{report.get('new_version')} "
                    f"({report.get('count')} файлов), backup={report.get('backup')}"
                )
                top.destroy()
                if messagebox.askyesno(
                    "Готово",
                    f"Установлена версия {report.get('new_version')}.\n"
                    f"Файлов: {report.get('count')}.\n\n"
                    "Перезапустить приложение сейчас?",
                ):
                    self._restart_app()
            except Exception as e:
                self.log_message(traceback.format_exc(), level="ERROR")
                messagebox.showerror("Обновление", str(e), parent=top)
            finally:
                self._busy(False)
                self.status_var.set("Готов к работе")

        bf = tk.Frame(top)
        bf.pack(pady=10)
        _btn(bf, "Установить", do_install, side=tk.LEFT, padx=4)
        _btn(bf, "Позже", top.destroy, side=tk.LEFT, padx=4)
        if info.html_url:
            def open_page():
                try:
                    if sys.platform == "darwin":
                        subprocess.Popen(["open", info.html_url])
                    elif sys.platform.startswith("win"):
                        os.startfile(info.html_url)  # type: ignore
                    else:
                        subprocess.Popen(["xdg-open", info.html_url])
                except Exception as e:
                    messagebox.showerror("Ошибка", str(e), parent=top)

            _btn(bf, "Открыть на GitHub", open_page, side=tk.LEFT, padx=4)

    def _restart_app(self):
        try:
            self._persist_settings()
        except Exception:
            pass
        if getattr(sys, "frozen", False):
            cmd = [sys.executable]
            cwd = str(APP_DIR)
        else:
            cmd = [sys.executable, str(APP_DIR / "app_desktop.py")]
            cwd = str(APP_DIR)
        try:
            subprocess.Popen(cmd, cwd=cwd)
        except Exception as e:
            messagebox.showerror(
                "Перезапуск",
                f"Не удалось перезапустить автоматически:\n{e}\n\n"
                "Закройте программу и запустите снова.",
            )
            return
        self.root.destroy()

    def _apply_saved_settings(self):
        s = load_settings(APP_DIR)
        if not s:
            return
        if s.get("summary_path"):
            self.summary_path.set(str(s["summary_path"]))
        if s.get("department"):
            self.dept_var.set(str(s["department"]))
        if s.get("year"):
            self.year_var.set(str(s["year"]))
            try:
                self.summary_cfg["year"] = int(s["year"])
            except (TypeError, ValueError):
                pass
        if "hide_zeros" in s:
            self.hide_zeros.set(bool(s["hide_zeros"]))
        if "filter_enabled" in s:
            self.filter_enabled.set(bool(s["filter_enabled"]))
        if s.get("start_date"):
            self.start_date_var.set(str(s["start_date"]))
        if s.get("end_date"):
            self.end_date_var.set(str(s["end_date"]))
        if "write_weeks" in s:
            self.write_weeks_var.set(bool(s["write_weeks"]))
        if "write_form" in s:
            self.write_form_var.set(bool(s["write_form"]))
        if s.get("plan_mode") in ("template", "emk"):
            self.plan_mode.set(s["plan_mode"])

    def _persist_settings(self):
        save_settings(
            APP_DIR,
            {
                "summary_path": self.summary_path.get(),
                "department": self.dept_var.get(),
                "year": self.year_var.get(),
                "hide_zeros": bool(self.hide_zeros.get()),
                "filter_enabled": bool(self.filter_enabled.get()),
                "start_date": self.start_date_var.get(),
                "end_date": self.end_date_var.get(),
                "write_weeks": bool(self.write_weeks_var.get()),
                "write_form": bool(self.write_form_var.get()),
                "plan_mode": self.plan_mode.get(),
            },
        )

    def _on_close(self):
        try:
            self._persist_settings()
        except Exception:
            pass
        self.root.destroy()

    def _open_log_file(self):
        path = APP_DIR / "analysis.log"
        if not path.exists():
            messagebox.showinfo("Журнал", "Файл analysis.log ещё не создан")
            return
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            elif sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _clear_log_file(self):
        if not messagebox.askyesno("Журнал", "Очистить analysis.log?"):
            return
        APP_LOG.clear()
        self.log_text.delete("1.0", tk.END)
        self.log_message("Журнал очищен")

    def export_unclassified(self):
        ops = self.get_view_ops()
        if ops.empty:
            messagebox.showwarning("Нет данных", "Нет операций")
            return
        uncl = ops[ops["Категория"] == "Не классифицировано"]
        if uncl.empty:
            messagebox.showinfo("Экспорт", "Неклассифицированных операций нет")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile="неклассифицировано.csv",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")],
        )
        if not path:
            return
        cols = [c for c in ("Дата", "КВС", "Код", "КСГ_название", "КСГ", "Услуга", "КСГ_подсказка") if c in uncl.columns]
        out = uncl[cols].copy()
        try:
            if path.lower().endswith(".xlsx"):
                out.to_excel(path, index=False)
            else:
                out.to_csv(path, index=False, encoding="utf-8-sig")
            self.log_message(f"Экспорт неклассифицированных: {path} ({len(out)} строк)")
            messagebox.showinfo("Готово", f"Сохранено {len(out)} строк:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    try:
        app = DesktopApp(root)
    except Exception:
        traceback.print_exc()
        messagebox.showerror("Старт", traceback.format_exc())
        root.destroy()
        raise
    root.mainloop()
