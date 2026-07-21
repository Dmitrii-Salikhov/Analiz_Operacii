# analyzers/dept_config.py
"""Конфигурация сводных и рубрикаторов по отделениям."""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List, Optional

DEFAULT_SHEET_NAMES = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь ",
    11: "Ноябрь",
    12: "Декабрь",
}

# Отчёты в «Отчеты других отделений/» для первичной инвентаризации
DEPT_REPORT_SOURCES: Dict[str, dict] = {
    "surg1": {
        "department": "1 Хирургическое отделение",
        "report_file": "1 хир.xlsx",
        "default_path": "1 хирургия операции сводная {year}.xlsx",
        "label": "1 хирургия",
    },
    "surg2": {
        "department": "2 Хирургическое отделение",
        "report_file": "2 хир.xlsx",
        "default_path": "2 хирургия операции сводная {year}.xlsx",
        "label": "2 хирургия",
    },
    "pedsurg": {
        "department": "Детское хирургическое отделение",
        "report_file": "дет хир.xlsx",
        "default_path": "детская хирургия операции сводная {year}.xlsx",
        "label": "детская хирургия",
    },
    "traum": {
        "department": "Травматологическое отделение",
        "report_file": "Травма.xlsx",
        "default_path": "травматология операции сводная {year}.xlsx",
        "label": "травматология",
    },
}


def dept_summary_key(config: dict, department_name: Optional[str] = None) -> str:
    """summary_key по имени отделения из department_profiles."""
    ensure_multi_dept_config(config)
    dept = department_name or (config.get("departments") or {}).get("main") or ""
    profiles = config.get("department_profiles") or {}
    prof = profiles.get(dept) or {}
    return str(prof.get("summary_key") or "lor")


def ensure_multi_dept_config(config: dict) -> dict:
    """
    Поднимает summaries / surgery_categories_by_dept из legacy summary / surgery_categories.
    Дублирует lor обратно в legacy-ключи для совместимости.
    """
    if not isinstance(config, dict):
        return config
    summaries = config.setdefault("summaries", {})
    by_dept = config.setdefault("surgery_categories_by_dept", {})

    legacy_summary = config.get("summary")
    if legacy_summary and "lor" not in summaries:
        summaries["lor"] = deepcopy(legacy_summary)

    legacy_cats = config.get("surgery_categories")
    if legacy_cats and "lor" not in by_dept:
        by_dept["lor"] = deepcopy(legacy_cats)

    if "lor" in summaries:
        config["summary"] = summaries["lor"]
    if "lor" in by_dept:
        config["surgery_categories"] = by_dept["lor"]

    return config


def get_summary_cfg(
    config: dict,
    *,
    summary_key: Optional[str] = None,
    department_name: Optional[str] = None,
) -> dict:
    ensure_multi_dept_config(config)
    key = summary_key or dept_summary_key(config, department_name)
    summaries = config.get("summaries") or {}
    if key in summaries:
        return summaries[key]
    if key == "lor" and config.get("summary"):
        return config["summary"]
    return {}


def get_surgery_categories(
    config: dict,
    *,
    summary_key: Optional[str] = None,
    department_name: Optional[str] = None,
) -> List[dict]:
    ensure_multi_dept_config(config)
    key = summary_key or dept_summary_key(config, department_name)
    by_dept = config.get("surgery_categories_by_dept") or {}
    if key in by_dept:
        return list(by_dept[key])
    if key == "lor" and config.get("surgery_categories"):
        return list(config["surgery_categories"])
    return []


def set_summary_cfg(config: dict, summary_key: str, summary_cfg: dict) -> None:
    ensure_multi_dept_config(config)
    config.setdefault("summaries", {})[summary_key] = summary_cfg
    if summary_key == "lor":
        config["summary"] = summary_cfg


def set_surgery_categories(config: dict, summary_key: str, categories: List[dict]) -> None:
    ensure_multi_dept_config(config)
    config.setdefault("surgery_categories_by_dept", {})[summary_key] = categories
    if summary_key == "lor":
        config["surgery_categories"] = categories


def form_4001_enabled(summary_cfg: dict) -> bool:
    form = (summary_cfg or {}).get("form_4001") or {}
    return bool(form.get("enabled", False))


def default_summary_filename(config: dict, summary_key: str, year: Optional[int] = None) -> str:
    cfg = get_summary_cfg(config, summary_key=summary_key)
    if cfg.get("default_path"):
        path = str(cfg["default_path"])
        y = year or int(cfg.get("year") or 2026)
        if "{year}" in path:
            return path.format(year=y)
        return path
    meta = DEPT_REPORT_SOURCES.get(summary_key) or {}
    tpl = meta.get("default_path") or "Операции сводная {year}.xlsx"
    y = year or int(cfg.get("year") or 2026)
    return tpl.format(year=y)


def default_sheet_names(summary_cfg: dict) -> Dict[int, str]:
    raw = (summary_cfg or {}).get("sheet_names") or DEFAULT_SHEET_NAMES
    return {int(k): str(v) for k, v in raw.items()}


def is_lor_department(config: dict, department_name: Optional[str] = None) -> bool:
    return dept_summary_key(config, department_name) == "lor"


def dept_display_label(summary_key: str) -> str:
    meta = DEPT_REPORT_SOURCES.get(summary_key) or {}
    return str(meta.get("label") or summary_key)
