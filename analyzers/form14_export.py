# analyzers/form14_export.py
"""Сбор строк маппинга ФСН 14 и запись/чтение Excel-черновика."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd
import yaml

from analyzers.dept_config import dept_full_name, ensure_multi_dept_config
from analyzers.form14_map import FORM14_LINES, map_code_auto, map_code_to_form14
from analyzers.form14_overrides import (
    COMMENT_COL,
    HISTOLOGY_COL,
    MANUAL_COL,
    histology_for,
    lookup_override,
    sanitize_dataframe_for_excel,
)
from analyzers.ksg_catalog import get_catalog

DEPT_KEYS = ("lor", "surg1", "surg2", "pedsurg", "traum")
DEFAULT_XLSX = "маппинг_ФСН14_4000_4001.xlsx"


def form14_line_choices() -> List[str]:
    """Значения для Combobox: '5.1 — из них — на ухе'."""
    return [f"{k} — {v}" for k, v in FORM14_LINES.items()]


def parse_line_choice(text: str) -> str:
    s = str(text or "").strip()
    if "—" in s:
        return s.split("—", 1)[0].strip()
    if "-" in s and s[0].isdigit():
        return s.split("-", 1)[0].strip()
    return s


def _dept_name(cfg: dict, key: str) -> str:
    return dept_full_name(cfg, key)


def iter_dept_categories(cfg: dict) -> Dict[str, List[dict]]:
    ensure_multi_dept_config(cfg)
    by_dept = dict(cfg.get("surgery_categories_by_dept") or {})
    if "lor" not in by_dept and cfg.get("surgery_categories"):
        by_dept["lor"] = cfg["surgery_categories"]
    return by_dept


def build_mapping_rows(
    cfg: dict,
    *,
    overrides: Optional[dict] = None,
    dept_keys: Optional[Sequence[str]] = None,
) -> List[dict]:
    """Строки для конструктора / Excel."""
    by_dept = iter_dept_categories(cfg)
    catalog = get_catalog()
    store = overrides if overrides is not None else {}
    keys = list(dept_keys) if dept_keys is not None else list(DEPT_KEYS)
    rows: List[dict] = []

    for key in keys:
        cats = by_dept.get(key) or []
        if not cats:
            continue
        dept = _dept_name(cfg, key)
        seen = set()
        for cat in cats:
            cat_name = str(cat.get("category") or "")
            codes = cat.get("codes") or []
            items = codes if codes else [""]
            for code in items:
                code = str(code or "").strip()
                if code and code in seen:
                    continue
                if code:
                    seen.add(code)
                ksg_name = catalog.name_for(code) if code else cat_name
                auto = map_code_auto(code, ksg_name, category=cat_name, summary_key=key)
                final = map_code_to_form14(
                    code,
                    ksg_name,
                    category=cat_name,
                    summary_key=key,
                    overrides=store,
                )
                ov = lookup_override(store, code=code, category=cat_name) if store else None
                manual = ov[0] if ov else ""
                comment = (ov[1].get("comment") or "") if ov else ""
                cat_hist = bool(cat.get("histology", False))
                ov_hist = histology_for(store, code=code, category=cat_name) if store else None
                effective_hist = ov_hist if ov_hist is not None else cat_hist
                rows.append(
                    {
                        "summary_key": key,
                        "Отделение": dept,
                        "Код": code,
                        "Категория": cat_name,
                        "Наименование_КСГ": ksg_name or cat_name,
                        "Класс_A16": final.a16_class,
                        "Авто": auto.line,
                        "Авто_название": auto.line_name,
                        "Строка_ФСН14": final.line,
                        "Название_строки": final.line_name,
                        MANUAL_COL: manual,
                        COMMENT_COL: comment,
                        HISTOLOGY_COL: "да" if effective_hist else "нет",
                        "histology": effective_hist,
                        "histology_override": ov_hist,
                        "Уверенность": final.confidence,
                        "Правило": final.rule,
                        "Примечание": final.notes,
                    }
                )
    return rows


def write_form14_excel(
    path: Path | str,
    cfg: dict,
    *,
    overrides: Optional[dict] = None,
) -> Path:
    path = Path(path)
    store = overrides if overrides is not None else {}
    all_rows = build_mapping_rows(cfg, overrides=store)
    all_df = pd.DataFrame(all_rows)

    summary_rows = []
    disputed_rows = []
    by_dept: Dict[str, list] = {}
    for r in all_rows:
        by_dept.setdefault(r["summary_key"], []).append(r)
        if r.get("Уверенность") == "low" or str(r.get("Строка_ФСН14")) == "21":
            disputed_rows.append(r)

    for key, rows in by_dept.items():
        high = sum(1 for x in rows if x.get("Уверенность") == "high")
        medium = sum(1 for x in rows if x.get("Уверенность") == "medium")
        low = sum(1 for x in rows if x.get("Уверенность") == "low")
        s21 = sum(1 for x in rows if str(x.get("Строка_ФСН14")) == "21")
        summary_rows.append(
            {
                "summary_key": key,
                "Отделение": rows[0]["Отделение"] if rows else key,
                "Кодов_уник": len(rows),
                "high": high,
                "medium": medium,
                "low": low,
                "строка_21": s21,
                "overrides": sum(1 for x in rows if x.get(MANUAL_COL)),
            }
        )

    lines_df = pd.DataFrame([{"Строка": k, "Наименование": v} for k, v in FORM14_LINES.items()])
    sum_df = pd.DataFrame(summary_rows)
    disp_df = pd.DataFrame(disputed_rows)
    if not all_df.empty:
        pivot = (
            all_df.groupby(["Строка_ФСН14", "Название_строки"], dropna=False)
            .size()
            .reset_index(name="Кодов")
            .sort_values("Строка_ФСН14")
        )
    else:
        pivot = pd.DataFrame()

    legend_df = pd.DataFrame(
        {
            "Поле": [
                "Строка_ФСН14_ручная",
                "Комментарий",
                "Морфология",
                "Приоритет",
                "Конструктор",
            ],
            "Значение": [
                "Заполните для переназначения; затем Импорт в программе",
                "Пояснение хирурга",
                "да/нет — морфологическое исследование (гистология)",
                "override YAML > калибровка ЛОР (только lor) > авто (класс+слова)",
                "Файл → Конструктор ФСН 14…",
            ],
        }
    )
    # нейтрализация формул (=HYPERLINK / DDE) во всех текстовых ячейках экспорта
    sum_df = sanitize_dataframe_for_excel(sum_df)
    all_df = sanitize_dataframe_for_excel(all_df)
    disp_df = sanitize_dataframe_for_excel(disp_df)
    pivot = sanitize_dataframe_for_excel(pivot)
    lines_df = sanitize_dataframe_for_excel(lines_df)
    legend_df = sanitize_dataframe_for_excel(legend_df)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sum_df.to_excel(writer, sheet_name="Сводка", index=False)
        all_df.to_excel(writer, sheet_name="Все коды", index=False)
        disp_df.to_excel(writer, sheet_name="Спорные_low_и_21", index=False)
        pivot.to_excel(writer, sheet_name="По строкам ФСН", index=False)
        lines_df.to_excel(writer, sheet_name="Справочник строк", index=False)
        legend_df.to_excel(writer, sheet_name="Легенда", index=False)
    return path


def form14_preview_rows_from_ops(
    month_ops: pd.DataFrame,
    *,
    summary_key: str,
    pension_age: int = 60,
    hide_zeros: bool = True,
    overrides: Optional[dict] = None,
) -> List[dict]:
    """
    Превью формы № 14 по операциям месяца для отделения (не шаблон ЛОР 4001).
    Каждая операция → строка ФСН 14 через map_code_to_form14(summary_key=...).
    Морфология: override histology, иначе флаг операции/категории.
    """
    from analyzers.form14_map import FORM14_LINES, map_code_to_form14

    empty = {
        "total": 0,
        "age_0_14": 0,
        "age_under_1": 0,
        "age_15_17": 0,
        "histology": 0,
        "senior": 0,
    }
    if month_ops is None or getattr(month_ops, "empty", True):
        return []

    store = overrides if overrides is not None else {}
    buckets: Dict[str, dict] = {}
    for _, r in month_ops.iterrows():
        code = str(r.get("Код") or "").strip()
        cat = str(r.get("Категория") or "").strip()
        m = map_code_to_form14(
            code, cat, category=cat, summary_key=summary_key, overrides=store
        )
        line = str(m.line or "21")
        b = buckets.setdefault(line, dict(empty))
        b["total"] += 1
        try:
            age = float(r.get("Возраст"))
        except (TypeError, ValueError):
            age = None
        if age is not None:
            if age < 1:
                b["age_under_1"] += 1
            if age < 15:
                b["age_0_14"] += 1
            elif age <= 17:
                b["age_15_17"] += 1
            if age >= pension_age:
                b["senior"] += 1
        ov_hist = histology_for(store, code=code, category=cat) if store else None
        if ov_hist is not None:
            is_hist = bool(ov_hist)
        else:
            is_hist = bool(r.get("Гистология") or r.get("histology"))
        if is_hist:
            b["histology"] += 1

    rows: List[dict] = []
    for line, name in FORM14_LINES.items():
        if line == "1":
            continue
        b = buckets.get(line) or empty
        if hide_zeros and int(b["total"]) == 0:
            continue
        rows.append(
            {
                "name": name,
                "line": line,
                "total": int(b["total"]),
                "age_0_14": int(b["age_0_14"]),
                "age_under_1": int(b["age_under_1"]),
                "age_15_17": int(b["age_15_17"]),
                "histology": int(b["histology"]),
                "senior": int(b["senior"]),
            }
        )

    ages = pd.to_numeric(month_ops.get("Возраст"), errors="coerce")
    pension = int((ages >= pension_age).sum()) if ages is not None else 0
    rows.append(
        {
            "name": "Операции лицам пенсионного возраста",
            "line": "",
            "total": pension,
            "age_0_14": 0,
            "age_under_1": 0,
            "age_15_17": 0,
            "histology": 0,
            "senior": "",
        }
    )
    rows.append(
        {
            "name": "Всего операций",
            "line": "",
            "total": int(len(month_ops)),
            "age_0_14": int((ages < 15).sum()) if ages is not None else 0,
            "age_under_1": int((ages < 1).sum()) if ages is not None else 0,
            "age_15_17": int(((ages >= 15) & (ages <= 17)).sum()) if ages is not None else 0,
            "histology": "",
            "senior": pension,
        }
    )
    return rows


def load_config(app_dir: Path | str) -> dict:
    p = Path(app_dir) / "config.yaml"
    with p.open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}
