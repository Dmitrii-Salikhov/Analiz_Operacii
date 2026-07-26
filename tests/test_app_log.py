# tests/test_app_log.py
"""Журнал analysis.log: лимит 500 строк и удаление старых."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.app_log import LOG_MAX_LINES, AppLog


def test_log_max_constant():
    assert LOG_MAX_LINES == 500


def test_append_trims_to_500(tmp_path: Path):
    path = tmp_path / "analysis.log"
    log = AppLog(path, max_lines=500)
    for i in range(520):
        log.append(f"msg-{i}", level="INFO")
    lines = log.read_lines(trim=False)
    assert len(lines) == 500
    assert "msg-19" not in lines[0]  # первые 20 удалены
    assert "msg-520" not in "\n".join(lines)
    assert "msg-519" in lines[-1]


def test_read_lines_trim_rewrites_file(tmp_path: Path):
    path = tmp_path / "analysis.log"
    # руками пишем больше лимита
    path.write_text("\n".join(f"old-{i}" for i in range(600)) + "\n", encoding="utf-8")
    log = AppLog(path, max_lines=500)
    lines = log.read_lines(trim=True)
    assert len(lines) == 500
    assert lines[0] == "old-100"
    assert lines[-1] == "old-599"
    # файл на диске тоже обрезан
    on_disk = [ln for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]
    assert len(on_disk) == 500


def test_clear(tmp_path: Path):
    path = tmp_path / "analysis.log"
    log = AppLog(path)
    log.append("x")
    log.clear()
    assert log.read_lines() == []
