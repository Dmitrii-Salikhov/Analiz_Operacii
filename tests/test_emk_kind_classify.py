# tests/test_emk_kind_classify.py
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.emk_kind_classify import classify_categories_by_emk
from analyzers.dept_inventory import category_display_name, shorten_category_label, strip_service_code_prefix


def test_strip_code_prefix():
    assert strip_service_code_prefix("A16.30.006 - ЛАПАРОТОМИЯ") == "ЛАПАРОТОМИЯ"
    assert strip_service_code_prefix("Некрэктомия") == "Некрэктомия"


def test_shorten_long_skin_excision():
    long = "Широкое иссечение новообразования кожи с реконструктивно-пластическим компонентом"
    short = shorten_category_label(long)
    assert short == "Иссечение новообразования кожи с реконструктивно-пластическим"
    assert "A16" not in short


def test_category_display_no_codes():
    used = set()
    n1 = category_display_name(ksg_name="Некрэктомия", code="A16.01.003", used=used)
    n2 = category_display_name(
        ksg_name="", sample="A16.30.006 - ЛАПАРОТОМИЯ", code="A16.30.006", used=used
    )
    assert n1 == "Некрэктомия"
    assert "A16" not in n2
    assert n2 == "ЛАПАРОТОМИЯ"


def test_classify_clear_and_disputed():
    ops = pd.DataFrame(
        [
            {"Категория": "Флегмона", "Тип_ЭМК": "экстренная"},
            {"Категория": "Флегмона", "Тип_ЭМК": "экстренная"},
            {"Категория": "Геморрой", "Тип_ЭМК": "плановая"},
            {"Категория": "Смешанная", "Тип_ЭМК": "плановая"},
            {"Категория": "Смешанная", "Тип_ЭМК": "экстренная"},
            {"Категория": "Без связи", "Тип_ЭМК": ""},
        ]
    )
    kind = classify_categories_by_emk(
        ops, category_names=["Флегмона", "Геморрой", "Смешанная", "Без связи"]
    )
    assert "Флегмона" in kind["emergency"]
    assert "Геморрой" in kind["plan"]
    assert "Без связи" in kind["no_emk"]
    assert len(kind["disputed"]) == 1
    assert kind["disputed"][0]["Категория"] == "Смешанная"
