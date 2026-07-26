# tests/test_flet_schemas.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from ui_flet.schema_loader import (
    actions_checklist,
    load_registry,
    load_schema,
    screens_by_status,
)


def test_registry_has_form14_done():
    reg = load_registry()
    screens = {s["id"]: s for s in reg.get("screens") or []}
    assert "form14" in screens
    assert screens["form14"]["status"] == "done"
    actions = {a["id"]: a for a in actions_checklist(reg)}
    assert actions["form14_ctor"]["status"] == "done"


def test_all_critical_actions_listed():
    required = {
        "load_surg",
        "load_emk",
        "write_excel",
        "form14_ctor",
        "add_category",
        "delete_category",
        "classify_emk_kind",
        "inventory",
        "check_updates",
    }
    ids = {a["id"] for a in actions_checklist()}
    assert required <= ids


def test_form14_schema_parameters():
    sch = load_schema("form14_constructor.schema.yaml")
    ids = {p["id"] for p in sch.get("parameters") or []}
    assert "form14_line" in ids
    assert "comment" in ids
    assert sch.get("storage") == "form14_overrides.yaml"
    assert "assign" in {a["id"] for a in sch.get("actions") or []}


def test_ui_settings_schema_keys():
    sch = load_schema("ui_settings.schema.yaml")
    ids = {p["id"] for p in sch.get("parameters") or []}
    for key in (
        "summary_path",
        "department",
        "year",
        "plan_mode",
        "theme",
        "summary_paths_by_dept",
        "write_weeks",
        "write_form",
    ):
        assert key in ids, key


def test_forced_emergency_schema_rules():
    sch = load_schema("forced_emergency.schema.yaml")
    rules = {r["id"] for r in sch.get("rules") or []}
    assert "drain_abscess" in rules
    assert "open_phlegmon_or_abscess" in rules


def test_screens_status_buckets():
    buckets = screens_by_status()
    assert any(s["id"] == "form14" for s in buckets["done"])
    assert buckets["shell"]  # ещё не всё портировано


def test_inventory_doc_exists():
    assert (APP / "docs" / "INVENTORY_FLET.md").exists()
