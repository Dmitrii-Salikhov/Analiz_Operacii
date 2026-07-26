# tests/test_form14_map.py
from __future__ import annotations

import sys
from pathlib import Path

import yaml

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.form_4001 import resolve_line_total_cats
from analyzers.form14_map import map_categories, map_code_to_form14


def test_tonsil_to_5_2():
    m = map_code_to_form14("A16.08.001.001", "Тонзилэктомия")
    assert m.line == "5.2"
    assert m.confidence == "high"


def test_ear_to_5_1():
    m = map_code_to_form14("A16.25.011", "Миринготомия")
    assert m.line == "5.1"


def test_appendix_to_9_2():
    m = map_code_to_form14("A16.18.009", "Аппендэктомия")
    assert m.line == "9.2"


def test_hernia_to_9_3():
    m = map_code_to_form14("A16.30.001", "Оперативное лечение пахово-бедренной грыжи")
    assert m.line == "9.3"


def test_skin_class_to_17():
    m = map_code_to_form14("A16.01.004", "Хирургическая обработка раны")
    assert m.line == "17"


def test_bone_to_15():
    m = map_code_to_form14("A16.03.022", "Остеосинтез")
    assert m.line == "15"
    assert m.confidence in ("high", "medium")


def test_lor_septoplasty_calibrated_to_6():
    """В сводной ЛОР нос/перегородка идут в стр. 6 (формулы), не в 5."""
    m = map_code_to_form14("A16.08.013.001", "Септопластика", category="Септопластика")
    assert m.line == "6"
    assert "калибровка" in m.rule


def test_lor_categories_match_formulas():
    cfg = yaml.safe_load((APP / "config.yaml").read_text(encoding="utf-8"))
    line_cats = resolve_line_total_cats(cfg["summary"]["form_4001"])
    expected = {}
    for line in ("5.1", "5.2", "6", "17", "6.1"):
        for c in line_cats.get(line, []):
            expected[c.strip()] = line

    for m in map_categories(cfg["surgery_categories"]):
        assert "категория: " in (m.notes or ""), m
        cat = m.notes.split("категория: ", 1)[1].strip()
        assert cat in expected, f"неизвестная категория для {m.code}: {cat!r}"
        assert m.line == expected[cat], f"{cat}: ожидали {expected[cat]}, получили {m.line}"
