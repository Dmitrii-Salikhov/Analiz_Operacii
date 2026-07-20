# tests/test_release_notes.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.release_notes import format_whats_new, notes_for_version, parse_release_sections


SAMPLE = """# Notes

## 1.0.4
- Новая фича A
- Новая фича B

## 1.0.3
- Старое C

## 1.0.2
- Ещё старше
"""


def test_parse_sections():
    s = parse_release_sections(SAMPLE)
    assert "1.0.4" in s and "1.0.3" in s
    assert "Новая фича A" in s["1.0.4"]
    assert "Старое C" in s["1.0.3"]


def test_notes_for_version(tmp_path: Path):
    p = tmp_path / "RELEASE_NOTES.md"
    p.write_text(SAMPLE, encoding="utf-8")
    assert "Новая фича A" in notes_for_version("1.0.4", path=p)
    assert notes_for_version("9.9.9", path=p) == ""


def test_format_whats_new_range(tmp_path: Path):
    p = tmp_path / "RELEASE_NOTES.md"
    p.write_text(SAMPLE, encoding="utf-8")
    text = format_whats_new("1.0.4", path=p, previous_version="1.0.3")
    assert "1.0.4" in text
    assert "Новая фича A" in text
    assert "Старое C" not in text

    text2 = format_whats_new("1.0.4", path=p, previous_version="1.0.2")
    assert "Новая фича A" in text2
    assert "Старое C" in text2
