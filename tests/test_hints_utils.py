"""Тесты бэкапов и проверки записи."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.backup_utils import (
    DEFAULT_BACKUP_KEEP,
    backups_dir,
    list_backups,
    make_backup,
    rotate_backups,
)
from analyzers.write_verify import format_verify_message, verify_write_report


def test_make_backup_goes_to_backups_folder(tmp_path):
    summary = tmp_path / "Операции сводная 2026.xlsx"
    summary.write_bytes(b"PK\x03\x04fake")
    bak = make_backup(summary, keep=20)
    assert bak is not None
    assert bak.parent == backups_dir(summary)
    assert bak.parent.name == "backups"
    assert bak.exists()
    assert bak.parent.exists()


def test_rotate_backups_keep_20(tmp_path):
    summary = tmp_path / "Операции сводная 2026.xlsx"
    summary.write_bytes(b"PK\x03\x04fake")
    bdir = backups_dir(summary)
    bdir.mkdir()
    for i in range(25):
        bak = bdir / f"Операции сводная 2026.202601{i:02d}_120000.bak.xlsx"
        bak.write_text(f"bak{i}")
    removed = rotate_backups(summary, keep=DEFAULT_BACKUP_KEEP)
    assert len(removed) == 5
    left = list_backups(summary)
    assert len(left) == 20


def test_list_includes_legacy_next_to_file(tmp_path):
    summary = tmp_path / "Операции сводная 2026.xlsx"
    summary.write_bytes(b"PK\x03\x04fake")
    legacy = tmp_path / "Операции сводная 2026.20260101_120000.bak.xlsx"
    legacy.write_text("old")
    found = list_backups(summary)
    assert legacy in found


def test_verify_empty_report(tmp_path):
    f = tmp_path / "t.xlsx"
    import openpyxl

    wb = openpyxl.Workbook()
    wb.save(f)
    wb.close()
    r = verify_write_report(f, {"months": {}})
    assert r.get("ok") is True
    assert "нечего" in (r.get("note") or "") or r.get("checked") == 0
    assert "OK" in format_verify_message({"ok": True, "checked": 3}) or "совпал" in format_verify_message(
        {"ok": True, "checked": 3}
    )
