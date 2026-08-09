# tests/test_app_paths.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.app_paths import resolve_app_dir


def test_resolve_app_dir_dev():
    root = resolve_app_dir(package_file=APP / "ui_flet" / "session.py")
    assert root == APP
    assert (root / "config.yaml").exists() or (root / "VERSION").exists()


def test_schema_root_matches_app():
    from ui_flet.schema_loader import APP_ROOT, SCHEMAS

    assert APP_ROOT == APP
    assert SCHEMAS == APP / "schemas"
    assert (SCHEMAS / "app_registry.yaml").exists()
