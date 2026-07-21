# tests/test_dept_config.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.dept_config import (
    default_summary_filename,
    dept_summary_key,
    ensure_multi_dept_config,
    form_4001_enabled,
    get_summary_cfg,
    get_surgery_categories,
)
from analyzers.dept_inventory import build_categories_from_codes, enrich_with_ksg, extract_codes_from_report


def test_migrate_legacy_config():
    cfg = {
        "departments": {"main": "Оториноларингологическое отделение"},
        "department_profiles": {
            "Оториноларингологическое отделение": {"summary_key": "lor"},
            "1 Хирургическое отделение": {"summary_key": "surg1"},
        },
        "summary": {"default_path": "lor.xlsx", "form_4001": {"enabled": True}},
        "surgery_categories": [{"category": "A", "codes": ["X"]}],
    }
    ensure_multi_dept_config(cfg)
    assert "lor" in cfg["summaries"]
    assert cfg["summaries"]["lor"]["default_path"] == "lor.xlsx"
    assert get_surgery_categories(cfg, summary_key="lor")[0]["category"] == "A"


def test_dept_summary_key():
    cfg = {
        "department_profiles": {
            "1 Хирургическое отделение": {"summary_key": "surg1"},
        }
    }
    assert dept_summary_key(cfg, "1 Хирургическое отделение") == "surg1"


def test_default_summary_path_template():
    cfg = {
        "summaries": {
            "surg1": {"default_path": "1 хирургия операции сводная {year}.xlsx", "year": 2026},
        }
    }
    assert default_summary_filename(cfg, "surg1", 2026) == "1 хирургия операции сводная 2026.xlsx"


def test_form_4001_disabled_for_surgery():
    cfg = {"summaries": {"surg1": {}}}
    assert not form_4001_enabled(get_summary_cfg(cfg, summary_key="surg1"))


def test_inventory_one_code_one_row():
    reports = APP / "Отчеты других отделений"
    path = reports / "дет хир.xlsx"
    if not path.exists():
        return
    table = extract_codes_from_report(path, "Детское хирургическое отделение")
    table = enrich_with_ksg(table)
    cats, rows, meta = build_categories_from_codes(table)
    assert len(cats) == len(table)
    assert len(rows) == len(cats)
    assert len(meta.get("plan_categories") or []) == len(cats)
    assert not meta.get("emergency_categories")
