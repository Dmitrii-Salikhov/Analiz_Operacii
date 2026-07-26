# ui_flet/session.py
"""Доменное состояние приложения без UI — общий backend для Flet (и будущего Tk)."""
from __future__ import annotations

import os
import traceback
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import yaml

from analyzers.app_log import AppLog
from analyzers.backup_utils import list_backups, restore_backup
from analyzers.category_registry import save_config, shift_totals_rows_by_delta
from analyzers.dept_config import (
    dept_summary_key,
    ensure_multi_dept_config,
    form_4001_enabled,
    get_summary_cfg,
    get_surgery_categories,
)
from analyzers.emk_compare import compare_plan_emergency
from analyzers.emk_kind_classify import (
    apply_kind_to_summary_cfg,
    classify_categories_by_emk,
    format_kind_report,
)
from analyzers.emk_loader import emk_department_stats, read_emk_stationary_report
from analyzers.export_report import export_month_like_summary
from analyzers.file_lock import excel_file_locked
from analyzers.form_4001 import compute_form_4001, form_4001_preview_rows
from analyzers.form14_export import form14_preview_rows_from_ops
from analyzers.form14_overrides import default_path as form14_overrides_path
from analyzers.form14_overrides import load_overrides
from analyzers.io_utils import OperationsStore, read_table
from analyzers.problem_codes import build_problem_codes_table
from analyzers.summary_writer import MONTH_RU, SummaryWriter, compute_month_weeks, read_sheet_weeks
from analyzers.surgery import SurgeryAnalyzer, build_summary_tables
from analyzers.ui_settings import load_settings, save_settings
from analyzers.updater import read_local_version
from analyzers.write_verify import format_verify_message, verify_write_report

APP_DIR = Path(__file__).resolve().parents[1]


@dataclass
class PreviewBundle:
    month_label: str = ""
    week_headers: List[str] = field(default_factory=list)
    cat_rows: List[List[Any]] = field(default_factory=list)  # [cat, w1..wn, tot]
    tot_rows: List[List[Any]] = field(default_factory=list)
    form_rows: List[dict] = field(default_factory=list)
    form_kind: str = ""  # "4001" | "14" | ""
    info: str = ""


class AppSession:
    def __init__(self, app_dir: Optional[Path] = None, log: Optional[Callable[[str], None]] = None):
        self.app_dir = Path(app_dir or APP_DIR)
        self._ui_log = log
        self.app_log = AppLog(self.app_dir / "analysis.log", max_lines=500)
        self.version = read_local_version(self.app_dir)

        self.config = self._load_config()
        ensure_multi_dept_config(self.config)
        self.department = (self.config.get("departments") or {}).get("main") or ""
        self.summary_key = "lor"
        self.summary_cfg: dict = {}
        self.store = OperationsStore()
        self.df_emk = None
        self.emk_path: Optional[str] = None
        self.loaded_department: Optional[str] = None
        self.last_batch_span: Tuple[Any, Any] = (None, None)
        self.cat_table = None
        self.totals_df = None
        self.weeks: list = []
        self.last_emk_compare = None
        self.preview = PreviewBundle()
        self.unclassified_rows: List[dict] = []
        self.disputed_rows: List[dict] = []
        self.emk_mismatch_rows: List[dict] = []
        self.kpi: Dict[str, str] = {
            "ops": "—",
            "patients": "—",
            "plan": "—",
            "emerg": "—",
            "period": "—",
            "files": "0 / нет",
            "diff": "—",
        }
        self.status = "Готов к работе"
        self.month_label_to_num: Dict[str, int] = {}
        self.preview_month = ""
        self.year = int((get_summary_cfg(self.config, summary_key="lor") or {}).get("year") or 2026)
        self.summary_path = str(self.app_dir / "Операции сводная 2026.xlsx")
        self.plan_mode = "template"
        self.hide_zeros = False
        self.filter_enabled = False
        self.start_date = "01.01.2026"
        self.end_date = "31.12.2026"
        self.write_weeks = True
        self.write_form = True
        self.last_surg_dir = str(self.app_dir)
        self.last_emk_dir = str(self.app_dir)
        self.summary_paths_by_dept: Dict[str, str] = {}
        self.theme = "light"
        self._apply_saved_settings()
        self._sync_dept()

    def _load_config(self) -> dict:
        with (self.app_dir / "config.yaml").open(encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def log(self, msg: str, level: str = "INFO") -> None:
        try:
            line = self.app_log.append(msg, level=level)
        except Exception:
            line = msg
        if self._ui_log:
            try:
                self._ui_log(line)
            except Exception:
                pass

    def read_log(self) -> str:
        try:
            lines = self.app_log.read_lines(trim=True)
            return "\n".join(lines)
        except Exception:
            p = self.app_dir / "analysis.log"
            return p.read_text(encoding="utf-8", errors="replace") if p.exists() else ""

    def log_lines_list(self) -> List[str]:
        try:
            return self.app_log.read_lines(trim=True)
        except Exception:
            return []

    def clear_log(self) -> None:
        try:
            self.app_log.clear()
        except Exception:
            (self.app_dir / "analysis.log").write_text("", encoding="utf-8")

    def departments(self) -> List[str]:
        return list((self.config.get("departments") or {}).get("list") or [])

    def _surgery_categories(self) -> list:
        return get_surgery_categories(self.config, summary_key=self.summary_key)

    def _sync_dept(self) -> None:
        ensure_multi_dept_config(self.config)
        self.summary_key = dept_summary_key(self.config, self.department)
        self.summary_cfg = get_summary_cfg(self.config, summary_key=self.summary_key)
        if not form_4001_enabled(self.summary_cfg):
            self.write_form = False
        # путь сводной для отделения
        if self.department in self.summary_paths_by_dept:
            self.summary_path = self.summary_paths_by_dept[self.department]
        else:
            default = self.summary_cfg.get("default_path") or "Операции сводная 2026.xlsx"
            default = str(default).format(year=self.year)
            cand = self.app_dir / default
            self.summary_path = str(cand)

    def set_department(self, name: str) -> None:
        self.department = name
        self._sync_dept()
        self.persist()

    def persist(self) -> None:
        self.summary_paths_by_dept[self.department] = self.summary_path
        save_settings(
            self.app_dir,
            {
                "summary_path": self.summary_path,
                "department": self.department,
                "year": str(self.year),
                "hide_zeros": self.hide_zeros,
                "filter_enabled": self.filter_enabled,
                "start_date": self.start_date,
                "end_date": self.end_date,
                "write_weeks": self.write_weeks,
                "write_form": self.write_form,
                "plan_mode": self.plan_mode,
                "last_surg_dir": self.last_surg_dir,
                "last_emk_dir": self.last_emk_dir,
                "summary_paths_by_dept": self.summary_paths_by_dept,
                "theme": self.theme,
            },
        )

    def _apply_saved_settings(self) -> None:
        s = load_settings(self.app_dir) or {}
        if not s:
            return
        if s.get("summary_path"):
            self.summary_path = str(s["summary_path"])
        if s.get("department"):
            self.department = str(s["department"])
        if s.get("year"):
            try:
                self.year = int(s["year"])
            except (TypeError, ValueError):
                pass
        self.hide_zeros = bool(s.get("hide_zeros", False))
        self.filter_enabled = bool(s.get("filter_enabled", False))
        if s.get("start_date"):
            self.start_date = str(s["start_date"])
        if s.get("end_date"):
            self.end_date = str(s["end_date"])
        self.write_weeks = bool(s.get("write_weeks", True))
        self.write_form = bool(s.get("write_form", True))
        if s.get("plan_mode") in ("template", "emk"):
            self.plan_mode = s["plan_mode"]
        if s.get("last_surg_dir"):
            self.last_surg_dir = str(s["last_surg_dir"])
        if s.get("last_emk_dir"):
            self.last_emk_dir = str(s["last_emk_dir"])
        if isinstance(s.get("summary_paths_by_dept"), dict):
            self.summary_paths_by_dept = {
                str(k): str(v) for k, v in s["summary_paths_by_dept"].items() if v
            }
        if s.get("theme") in ("light", "dark"):
            self.theme = str(s["theme"])

    @staticmethod
    def _parse_date(text: str) -> datetime:
        return datetime.strptime(str(text).strip(), "%d.%m.%Y")

    def sources_text(self) -> str:
        if self.store.ops.empty:
            return "Нет загруженных журналов"
        self.store.refresh_source_meta()
        blocks = []
        for src in self.store.sources:
            meta = self.store.source_meta.get(src) or {}
            d0, d1, n = meta.get("date_from"), meta.get("date_to"), meta.get("count", 0)
            if d0 is not None and d1 is not None:
                blocks.append(
                    f"• {src}\n  {d0.strftime('%d.%m.%Y')} – {d1.strftime('%d.%m.%Y')}\n  операций: {n}"
                )
            else:
                blocks.append(f"• {src}\n  операций: {n}")
        return "\n\n".join(blocks)

    def emk_status_text(self) -> str:
        if self.df_emk is None:
            return "ЭМК: не загружен"
        name = os.path.basename(self.emk_path) if self.emk_path else "загружен"
        return f"ЭМК: {name} ({len(self.df_emk)} стр.)"

    def refresh_files_kpi(self) -> None:
        self.kpi["files"] = f"{len(self.store.sources)} / {'да' if self.df_emk is not None else 'нет'}"

    def get_view_ops(self) -> pd.DataFrame:
        ops = self.store.ops.copy()
        if ops.empty:
            return ops
        if self.filter_enabled:
            start = pd.Timestamp(self._parse_date(self.start_date))
            end = pd.Timestamp(self._parse_date(self.end_date)) + timedelta(days=1) - timedelta(seconds=1)
            dates = pd.to_datetime(ops["Дата"], errors="coerce")
            ops = ops.loc[(dates >= start) & (dates <= end)].copy()
        return ops

    def ingest_surg_paths(self, paths: List[str]) -> None:
        dept = self.department
        any_added = False
        for path in paths:
            try:
                df = read_table(path)
                n_all = len(df)
                analyzer = SurgeryAnalyzer(df, dept, self._surgery_categories(), emk_df=self.df_emk)
                n_dept = len(analyzer.df)
                self.log(f"{os.path.basename(path)}: фильтр «{dept}» — {n_dept} из {n_all} строк журнала")
                ops = analyzer.extract_operations()
                if ops.empty:
                    self.log(f"Нет операций: {os.path.basename(path)}")
                    continue
                info = self.store.add(ops, path)
                any_added = True
                self.last_batch_span = (info.get("date_from"), info.get("date_to"))
                msg = (
                    f"{os.path.basename(path)}: +{info['added']}, вытеснено {info['removed']}, "
                    f"всего {info['total']}"
                )
                if info.get("date_from") is not None:
                    msg += f" | {info['date_from'].strftime('%d.%m.%Y')}–{info['date_to'].strftime('%d.%m.%Y')}"
                self.log(msg)
                uncl = ops[ops["Категория"] == "Не классифицировано"]
                if len(uncl):
                    codes = sorted({c for c in uncl["Код"].dropna().unique() if c})
                    self.log(f"  не классифицировано: {len(uncl)} опер., коды: {codes}")
            except Exception as e:
                self.log(f"Ошибка {os.path.basename(path)}: {e}\n{traceback.format_exc()}")
                raise
        if any_added:
            self.loaded_department = dept
        self.refresh_files_kpi()
        self.run_analysis()

    def load_emk_path(self, path: str) -> None:
        self.last_emk_dir = str(Path(path).parent)
        if str(path).lower().endswith(".csv"):
            self.df_emk = read_table(path)
        else:
            self.df_emk = read_emk_stationary_report(path)
        self.emk_path = path
        self.refresh_files_kpi()
        stats = emk_department_stats(self.df_emk)
        dept = self.department
        dept_rows = stats.get(dept)
        if dept_rows is None:
            dept_rows = sum(v for k, v in stats.items() if dept and dept in str(k))
        extra = f", отделение «{dept}»: {dept_rows} стр." if dept_rows else ""
        self.log(f"ЭМК: {path} ({len(self.df_emk)} строк{extra})")
        if not self.store.ops.empty:
            self.run_analysis()
        self.persist()

    def clear_store(self) -> None:
        self.store.clear()
        self.cat_table = self.totals_df = None
        self.weeks = []
        self.loaded_department = None
        self.last_emk_compare = None
        self.preview = PreviewBundle()
        self.unclassified_rows = []
        self.disputed_rows = []
        self.emk_mismatch_rows = []
        for k in self.kpi:
            self.kpi[k] = "—" if k != "files" else "0 / нет"
        self.refresh_files_kpi()
        self.status = "Очищено"
        self.log("Накопитель очищен")

    def _weeks_for_month(self, year: int, month: int):
        path = self.summary_path.strip()
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

    def _update_month_choices(self, ops: pd.DataFrame) -> None:
        months = sorted({int(m) for m in pd.to_datetime(ops["Дата"]).dt.month.dropna().unique()})
        self.month_label_to_num = {}
        labels = []
        for m in months:
            label = MONTH_RU.get(m, str(m))
            self.month_label_to_num[label] = m
            labels.append(label)
        if labels and self.preview_month not in self.month_label_to_num:
            self.preview_month = labels[-1]

    def run_analysis(self) -> None:
        ops = self.get_view_ops()
        if ops.empty:
            self.status = "Нет операций"
            return
        cat_table, totals_df, weeks = build_summary_tables(
            ops, self.summary_cfg, self._surgery_categories()
        )
        self.cat_table, self.totals_df, self.weeks = cat_table, totals_df, weeks
        self._update_month_choices(ops)
        self.build_preview()
        self._fill_unclassified(ops)
        self._fill_disputed(ops)
        self._update_kpis(ops, totals_df)
        self.refresh_emk_compare(select=False)
        self.status = f"Готово: {len(ops)} операций, {len(weeks)} нед."
        self.log(self.status)

    def build_preview(self) -> None:
        ops = self.get_view_ops()
        pb = PreviewBundle()
        if ops.empty or not self.preview_month or self.preview_month not in self.month_label_to_num:
            self.preview = pb
            return
        month = int(self.month_label_to_num[self.preview_month])
        weeks = self._weeks_for_month(self.year, month) or compute_month_weeks(self.year, month)
        week_headers = [f"{s.strftime('%d.%m')}-{e.strftime('%d.%m')}" for s, e in weeks]
        month_ops = ops[pd.to_datetime(ops["Дата"]).dt.month == month].copy()
        cat_order = list(self.summary_cfg.get("category_rows", {}).keys())
        counts_map = {cat: [0] * len(weeks) for cat in cat_order}
        for _, r in month_ops.iterrows():
            cat = r["Категория"]
            if cat not in counts_map:
                continue
            d = pd.Timestamp(r["Дата"]).date()
            for wi, (s, e) in enumerate(weeks):
                if s <= d <= e:
                    counts_map[cat][wi] += 1
                    break
        cat_rows = []
        for cat in cat_order:
            counts = counts_map[cat]
            total = sum(counts)
            if self.hide_zeros and total == 0:
                continue
            cat_rows.append([cat] + counts + [total])

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
        tot_rows = []
        for name, arr in (
            ("Всего операций", arrays["ops"]),
            ("Экстренно операций", arrays["emerg"]),
            ("План операций", arrays["plan"]),
            ("Дети всего", arrays["kids"]),
            ("Взрослые", adults),
            ("Человек", arrays["people"]),
        ):
            tot_rows.append([name] + arr + [sum(arr)])

        form_cfg = self.summary_cfg.get("form_4001") or {}
        pension_age = int(self.config.get("thresholds", {}).get("pension_age", 60))
        if form_4001_enabled(self.summary_cfg):
            stats = compute_form_4001(
                month_ops,
                self._surgery_categories(),
                pension_age=pension_age,
                form_cfg=form_cfg,
            )
            form_rows = form_4001_preview_rows(stats)
            form_kind = "4001"
        else:
            # не ЛОР: превью по форме № 14 из операций отделения (без шаблона ЛОР 5/5.1/5.2)
            ov = load_overrides(form14_overrides_path(self.app_dir))
            form_rows = form14_preview_rows_from_ops(
                month_ops,
                summary_key=self.summary_key,
                pension_age=pension_age,
                hide_zeros=True,
                overrides=ov,
            )
            form_kind = "14"
        pb.month_label = self.preview_month
        pb.week_headers = week_headers
        pb.cat_rows = cat_rows
        pb.tot_rows = tot_rows
        pb.form_rows = form_rows
        pb.form_kind = form_kind
        pb.info = (
            f"Превью: {self.preview_month} | операций месяца {len(month_ops)} | "
            f"недель {len(weeks)} | категории {len(cat_rows)}"
        )
        self.preview = pb

    def _fill_unclassified(self, ops: pd.DataFrame) -> None:
        uncl = ops[ops["Категория"] == "Не классифицировано"]
        rows = []
        for _, r in uncl.iterrows():
            dt = r["Дата"]
            dt_s = dt.strftime("%d.%m.%Y") if hasattr(dt, "strftime") else str(dt)
            rows.append(
                {
                    "Дата": dt_s,
                    "КВС": r.get("КВС"),
                    "Код": r.get("Код"),
                    "КСГ_название": r.get("КСГ_название", ""),
                    "КСГ": r.get("КСГ", ""),
                    "Услуга": str(r.get("Услуга", "") or "")[:80],
                }
            )
        self.unclassified_rows = rows

    def _fill_disputed(self, ops: pd.DataFrame) -> None:
        rows = []
        if ops is None or ops.empty or "Спор_ключей" not in ops.columns:
            self.disputed_rows = rows
            return
        disp = ops[ops["Спор_ключей"].fillna(False).astype(bool)]
        for _, r in disp.iterrows():
            dt = r["Дата"]
            dt_s = dt.strftime("%d.%m.%Y") if hasattr(dt, "strftime") else str(dt)
            rows.append(
                {
                    "Дата": dt_s,
                    "КВС": r.get("КВС"),
                    "Код": r.get("Код"),
                    "Услуга": str(r.get("Услуга", "") or "")[:100],
                    "Категория": r.get("Категория"),
                    "Кандидаты": r.get("Спорные_категории", ""),
                }
            )
        self.disputed_rows = rows

    def _update_kpis(self, ops: pd.DataFrame, totals_df) -> None:
        self.kpi["ops"] = str(len(ops))
        try:
            self.kpi["patients"] = str(int(ops["КВС"].nunique()))
        except Exception:
            self.kpi["patients"] = "—"
        emerg = set(self.summary_cfg.get("emergency_categories") or [])
        plan = set(self.summary_cfg.get("plan_categories") or [])
        n_e = int(ops["Категория"].isin(emerg).sum()) if emerg else 0
        n_p = int(ops["Категория"].isin(plan).sum()) if plan else 0
        n = max(len(ops), 1)
        self.kpi["emerg"] = f"{round(100 * n_e / n)}%"
        self.kpi["plan"] = f"{round(100 * n_p / n)}%"
        d0 = pd.to_datetime(ops["Дата"]).min()
        d1 = pd.to_datetime(ops["Дата"]).max()
        self.kpi["period"] = f"{d0.strftime('%d.%m')}–{d1.strftime('%d.%m.%Y')}"
        self.refresh_files_kpi()

    def refresh_emk_compare(self, select: bool = True) -> dict:
        ops = self.get_view_ops()
        if ops.empty or self.df_emk is None:
            self.emk_mismatch_rows = []
            self.kpi["diff"] = "—"
            return {}
        result = compare_plan_emergency(ops, self.summary_cfg, department=self.department)
        self.last_emk_compare = result
        mismatches = result.get("mismatches") or []
        rows = []
        for m in mismatches:
            dt = m.get("Дата")
            dt_s = dt.strftime("%d.%m.%Y") if hasattr(dt, "strftime") else str(dt or "")
            rows.append(
                {
                    "Дата": dt_s,
                    "КВС": m.get("КВС"),
                    "Категория": m.get("Категория"),
                    "Код": m.get("Код"),
                    "Шаблон": m.get("Шаблон"),
                    "ЭМК": m.get("ЭМК"),
                    "Диагноз": m.get("Диагноз"),
                    "Услуга": str(m.get("Услуга") or "")[:60],
                }
            )
        self.emk_mismatch_rows = rows
        self.kpi["diff"] = str(len(rows))
        return result

    def write_excel(self, *, write_weeks: bool, write_form: bool) -> dict:
        path = self.summary_path.strip()
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"Сводная не найдена: {path}")
        if excel_file_locked(path):
            raise RuntimeError(f"Файл занят в Excel:\n{path}")
        if self.store.ops.empty:
            raise ValueError("Нет операций")
        d_min, d_max = self.last_batch_span
        if d_min is None or d_max is None:
            d_min, d_max = self.store.date_span(self.store.ops)
        write_form = bool(write_form) and form_4001_enabled(self.summary_cfg)
        writer = SummaryWriter(
            path,
            self.summary_cfg,
            department=self.department,
            categories=self._surgery_categories(),
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
        blank_delta = int(report.get("blank_delta") or 0)
        if blank_delta:
            shift_totals_rows_by_delta(self.config, blank_delta, summary_key=self.summary_key)
            save_config(self.config, self.app_dir / "config.yaml")
            self._sync_dept()
        if write_weeks:
            try:
                vres = verify_write_report(path, report)
                report["verify_msg"] = format_verify_message(vres)
            except Exception as ve:
                report["verify_msg"] = f"Проверка записи не выполнена: {ve}"
        self.write_weeks = write_weeks
        self.write_form = write_form
        self.persist()
        self.log(
            f"Сводная обновлена: ячеек {report.get('cells_written', 0)}, "
            f"месяцы {list((report.get('months') or {}).keys())}"
        )
        self.status = f"Запись: {report.get('cells_written', 0)} ячеек"
        return report

    def export_simple_report(self, out_path: str) -> dict:
        path_tpl = self.summary_path.strip()
        if not path_tpl or not os.path.exists(path_tpl):
            raise FileNotFoundError(f"Нужен шаблон сводной:\n{path_tpl}")
        month = self.month_label_to_num.get(self.preview_month)
        report = export_month_like_summary(
            path_tpl,
            out_path,
            self.get_view_ops(),
            self.summary_cfg,
            department=self.department,
            categories=self._surgery_categories(),
            pension_age=int(self.config.get("thresholds", {}).get("pension_age", 60)),
            month=month,
            year=self.year,
        )
        self.log(f"Экспорт отчёта: {out_path}")
        return report

    def export_unclassified(self, out_path: str) -> int:
        df = pd.DataFrame(self.unclassified_rows)
        if out_path.lower().endswith(".xlsx"):
            df.to_excel(out_path, index=False)
        else:
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
        return len(df)

    def export_problem_codes(self, out_path: str) -> int:
        ops = self.get_view_ops()
        table = build_problem_codes_table(ops)
        if out_path.lower().endswith(".xlsx"):
            table.to_excel(out_path, index=False)
        else:
            table.to_csv(out_path, index=False, encoding="utf-8-sig")
        return len(table)

    def export_emk_mismatches(self, out_path: str) -> int:
        df = pd.DataFrame(self.emk_mismatch_rows)
        if out_path.lower().endswith(".xlsx"):
            df.to_excel(out_path, index=False)
        else:
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
        return len(df)

    def classify_kinds_from_emk(self) -> dict:
        if self.df_emk is None:
            raise ValueError("Сначала загрузите ЭМК")
        ops = self.get_view_ops()
        if ops.empty:
            raise ValueError("Нет операций")
        # пересчёт с ЭМК через analyzer уже в ops если загружали с emk;
        # классификация категорий по Тип_ЭМК
        names = list((self.summary_cfg.get("category_rows") or {}).keys())
        kind = classify_categories_by_emk(ops, category_names=names)
        updated = apply_kind_to_summary_cfg(dict(self.summary_cfg), kind)
        from analyzers.dept_config import set_summary_cfg

        set_summary_cfg(self.config, self.summary_key, updated)
        save_config(self.config, self.app_dir / "config.yaml")
        self._sync_dept()
        self.log(format_kind_report(kind))
        return kind

    def list_backups_for_summary(self) -> list:
        return list_backups(self.summary_path)

    def restore_backup_file(self, bak_path: str) -> str:
        return restore_backup(bak_path, self.summary_path)

    def form4001_enabled(self) -> bool:
        return form_4001_enabled(self.summary_cfg)

    def open_path(self, path: str) -> None:
        import subprocess
        import sys

        p = str(path)
        if sys.platform == "darwin":
            subprocess.Popen(["open", p])
        elif sys.platform.startswith("win"):
            os.startfile(p)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", p])
