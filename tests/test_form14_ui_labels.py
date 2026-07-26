# tests/test_form14_ui_labels.py
"""Подписи конструктора ФСН 14 и AppContext по отделениям."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from ui_flet.app_context import AppContext
from ui_flet.screens.form14 import CONF_RU, _ru_confidence, _ru_rule


def test_confidence_ru():
    assert _ru_confidence("high") == "высокая"
    assert _ru_confidence("medium") == "средняя"
    assert _ru_confidence("low") == "низкая"
    assert set(CONF_RU) == {"high", "medium", "low"}


def test_rule_ru_no_english_jargon():
    s = _ru_rule("keyword: аппендэктомия")
    assert "по словам:" in s
    assert "keyword" not in s
    s2 = _ru_rule("override YAML > калибровка")
    assert "ручная правка" in s2
    assert "override" not in s2.lower() or "ручная" in s2


def test_app_context_dept_dropdown_full_names():
    ctx = AppContext(APP)
    keys = ctx.dept_keys()
    assert "lor" in keys
    assert "surg1" in keys
    opts = ctx.dept_dropdown_options()
    assert len(opts) == len(keys)
    for key, title in opts:
        assert key in keys
        assert title == ctx.dept_full_name(key)
        assert title != key
        assert " " in title or "ое" in title.lower()
