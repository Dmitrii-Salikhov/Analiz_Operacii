"""Тесты бэкапов и проверки записи."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.backup_utils import list_backups, rotate_backups
from analyzers.write_verify import format_verify_message, verify_write_report


def test_rotate_backups(tmp_path):
    summary = tmp_path / "Операции сводная 2026.xlsx"
    summary.write_bytes(b"PK\x03\x04fake")
    created = []
    for i in range(5):
        bak = tmp_path / f"Операции сводная 2026.2026010{i}_120000.bak.xlsx"
        bak.write_text(f"bak{i}")
        created.append(bak)
    removed = rotate_backups(summary, keep=2)
    assert len(removed) == 3
    left = list_backups(summary)
    assert len(left) == 2


def test_verify_empty_report(tmp_path):
    f = tmp_path / "t.xlsx"
    # minimal fake — verify should handle missing months
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
