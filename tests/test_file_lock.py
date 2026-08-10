# tests/test_file_lock.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.file_lock import (
    FileLockedError,
    excel_file_locked,
    ensure_excel_writable,
    probe_excel_lock,
    remove_stale_excel_lock,
)


def test_missing_file_not_locked(tmp_path: Path):
    p = tmp_path / "no.xlsx"
    assert excel_file_locked(p) is False
    ensure_excel_writable(p)


def test_writable_file(tmp_path: Path):
    p = tmp_path / "ok.xlsx"
    p.write_bytes(b"PK")
    assert excel_file_locked(p) is False
    ensure_excel_writable(p)


def test_stale_lock_sibling_does_not_block(tmp_path: Path):
    """~$… остался после сбоя Excel, но сам файл открывается — не блокируем."""
    p = tmp_path / "book.xlsx"
    p.write_bytes(b"PK")
    lock = tmp_path / "~$book.xlsx"
    lock.write_text("lock")
    st = probe_excel_lock(p)
    assert st.locked is False
    assert st.reason == "stale_lock_only"
    assert excel_file_locked(p) is False
    ensure_excel_writable(p)
    assert remove_stale_excel_lock(p) is True
    assert not lock.exists()


def test_open_denied(tmp_path: Path, monkeypatch):
    p = tmp_path / "busy.xlsx"
    p.write_bytes(b"PK")

    def boom(*_a, **_k):
        raise PermissionError(13, "denied")

    monkeypatch.setattr("builtins.open", boom)
    assert excel_file_locked(p) is True
    try:
        ensure_excel_writable(p)
        assert False
    except FileLockedError as e:
        assert e.reason == "open_denied"


def test_install_file_rename_fallback(tmp_path: Path, monkeypatch):
    from analyzers import updater

    src = tmp_path / "new.bin"
    dest = tmp_path / "app.exe"
    src.write_bytes(b"NEW")
    dest.write_bytes(b"OLD")

    calls = {"n": 0}
    real_copy = updater.shutil.copy2

    def flaky_copy(s, d):
        calls["n"] += 1
        if calls["n"] <= 2 and Path(d) == dest:
            raise PermissionError(32, "busy")
        return real_copy(s, d)

    monkeypatch.setattr(updater.shutil, "copy2", flaky_copy)
    updater._install_file(src, dest)
    assert dest.read_bytes() == b"NEW"
