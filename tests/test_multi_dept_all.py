# tests/test_multi_dept_all.py
"""Параметризованные проверки по всем хирургическим отделениям и ключевым параметрам."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
import yaml

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.dept_config import (
    DEPT_REPORT_SOURCES,
    dept_full_name,
    dept_summary_key,
    ensure_multi_dept_config,
    form_4001_enabled,
    get_summary_cfg,
    get_surgery_categories,
)
from analyzers.dept_inventory import extract_codes_from_report
from analyzers.form14_export import DEPT_KEYS, build_mapping_rows, form14_line_choices, parse_line_choice
from analyzers.io_utils import read_table
from analyzers.surgery import SurgeryAnalyzer, build_summary_tables

REPORTS = APP / "Отчеты других отделений"

# (summary_key, department_name, report_relpath | None, summary_xlsx | None)
DEPT_CASES = [
    (
        "lor",
        "Оториноларингологическое отделение",
        None,  # журналы ЛОР в корне проекта
        "Операции сводная 2026.xlsx",
    ),
    (
        "surg1",
        "1 Хирургическое отделение",
        "1 хир.xlsx",
        "1 хирургия операции сводная 2026.xlsx",
    ),
    (
        "surg2",
        "2 Хирургическое отделение",
        "2 хир.xlsx",
        "2 хирургия операции сводная 2026.xlsx",
    ),
    (
        "pedsurg",
        "Детское хирургическое отделение",
        "дет хир.xlsx",
        "детская хирургия операции сводная 2026.xlsx",
    ),
    (
        "traum",
        "Травматологическое отделение",
        "Травма.xlsx",
        "травматология операции сводная 2026.xlsx",
    ),
]

LOR_JOURNALS = [
    APP / "Отчет по выполненным операциям и операционным столам (17).xlsx",
    APP / "Отчет по выполненным операциям и операционным столам (19).xlsx",
]


def _cfg():
    cfg = yaml.safe_load((APP / "config.yaml").read_text(encoding="utf-8"))
    ensure_multi_dept_config(cfg)
    return cfg


@pytest.fixture(scope="module")
def cfg():
    return _cfg()


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_dept_full_name_not_raw_key(cfg, key, dept_name, report, summary):
    name = dept_full_name(cfg, key)
    assert name == dept_name
    assert key not in name  # не «lor» / «surg1» в UI-названии
    assert len(name) > 5


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_summary_key_roundtrip(cfg, key, dept_name, report, summary):
    assert dept_summary_key(cfg, dept_name) == key


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_categories_and_form4001_flag(cfg, key, dept_name, report, summary):
    cats = get_surgery_categories(cfg, summary_key=key)
    assert len(cats) >= 10, f"{key}: слишком мало категорий ({len(cats)})"
    scfg = get_summary_cfg(cfg, summary_key=key)
    assert "category_rows" in scfg or "year" in scfg
    if key == "lor":
        assert form_4001_enabled(scfg) is True
    else:
        assert form_4001_enabled(scfg) is False


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_form14_mapping_rows_full_dept_name(cfg, key, dept_name, report, summary):
    rows = build_mapping_rows(cfg, overrides={}, dept_keys=[key])
    assert rows, f"{key}: пустой маппинг ФСН 14"
    for r in rows[:50]:
        assert r["summary_key"] == key
        assert r["Отделение"] == dept_name
        assert r.get("Строка_ФСН14")
        assert r.get("Уверенность") in ("high", "medium", "low")


def test_form14_all_dept_keys_covered(cfg):
    assert set(DEPT_KEYS) == {c[0] for c in DEPT_CASES}
    rows = build_mapping_rows(cfg, overrides={}, dept_keys=list(DEPT_KEYS))
    keys = {r["summary_key"] for r in rows}
    assert keys == set(DEPT_KEYS)
    names = {r["Отделение"] for r in rows}
    for _, dept_name, _, _ in DEPT_CASES:
        assert dept_name in names


def test_form14_line_choices_parse():
    choices = form14_line_choices()
    assert len(choices) >= 10
    assert parse_line_choice(choices[0]).replace(".", "").isdigit() or parse_line_choice(choices[0])[0].isdigit()
    assert parse_line_choice("9.2 — аппендэктомия") == "9.2"


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_journal_extract_and_summary_tables(cfg, key, dept_name, report, summary):
    cats = get_surgery_categories(cfg, summary_key=key)
    scfg = get_summary_cfg(cfg, summary_key=key)

    if key == "lor":
        journal = next((p for p in LOR_JOURNALS if p.exists()), None)
        if journal is None:
            pytest.skip("нет журнала ЛОР")
        df = read_table(str(journal))
    else:
        path = REPORTS / report
        if not path.exists():
            pytest.skip(f"нет отчёта {report}")
        # sanity: коды из журнала
        codes_df = extract_codes_from_report(path, dept_name)
        assert not codes_df.empty, f"{key}: в отчёте нет кодов A16"
        df = read_table(str(path))

    ops = SurgeryAnalyzer(df, dept_name, cats).extract_operations()
    assert not ops.empty, f"{key}: 0 операций после классификации"
    assert "Категория" in ops.columns
    cat_table, totals_df, weeks = build_summary_tables(ops, scfg, cats)
    assert cat_table is not None
    assert totals_df is not None
    assert len(weeks) >= 1
    # хотя бы одна категория с ненулевым итогом
    if hasattr(cat_table, "empty"):
        assert not cat_table.empty


@pytest.mark.parametrize("key,dept_name,report,summary", DEPT_CASES, ids=[c[0] for c in DEPT_CASES])
def test_summary_workbook_exists(key, dept_name, report, summary):
    path = APP / summary
    if not path.exists():
        pytest.skip(f"нет сводной {summary}")
    assert path.stat().st_size > 1000


def test_dept_report_sources_meta_matches_cases():
    for key, dept_name, report, _ in DEPT_CASES:
        if key == "lor":
            continue
        meta = DEPT_REPORT_SOURCES[key]
        assert meta["department"] == dept_name
        assert meta["report_file"] == report


def test_non_lor_ignores_lor_code_calibration():
    """Код из MANUAL_LINE_BY_CODE ЛОР не должен задавать строку для хирургии."""
    from analyzers.form14_map import map_code_to_form14

    code = "A16.01.011"
    cat = "Вскрытие фурункула (карбункула)"
    m_surg = map_code_to_form14(code, cat, category=cat, summary_key="surg1")
    m_lor = map_code_to_form14(code, cat, category=cat, summary_key="lor")
    assert m_surg.line == "17"
    assert "калибровка ЛОР" not in m_surg.rule
    assert "калибровка ЛОР" in m_lor.rule


def test_build_mapping_surg1_not_dominated_by_lor_rules():
    from collections import Counter

    from analyzers.form14_export import build_mapping_rows

    cfg = yaml.safe_load((APP / "config.yaml").read_text(encoding="utf-8"))
    rows = build_mapping_rows(cfg, overrides={}, dept_keys=["surg1"])
    assert rows
    lor_rules = [r for r in rows if "калибровка ЛОР" in str(r.get("Правило") or "")]
    assert not lor_rules, f"неожиданная калибровка ЛОР: {lor_rules[:3]}"
    lines = Counter(r["Строка_ФСН14"] for r in rows)
    lor_share = sum(v for k, v in lines.items() if str(k).startswith("5"))
    assert lor_share < len(rows) * 0.15


def test_form14_preview_from_ops_not_lor_template():
    """Для surg1 превью формы № 14 — по операциям, без строк ЛОР 5/5.1/5.2 при нулях."""
    import pandas as pd

    from analyzers.form14_export import form14_preview_rows_from_ops

    ops = pd.DataFrame(
        [
            {"Код": "A16.18.009", "Категория": "Аппендэктомия", "Возраст": 40, "Гистология": False},
            {"Код": "A16.30.001", "Категория": "Грыжесечение", "Возраст": 55, "Гистология": False},
            {"Код": "A16.01.004", "Категория": "ПХО", "Возраст": 30, "Гистология": False},
        ]
    )
    rows = form14_preview_rows_from_ops(ops, summary_key="surg1", hide_zeros=True)
    lines = [r["line"] for r in rows if r.get("line")]
    assert "9.2" in lines or "9.3" in lines or "9" in lines
    assert "5.1" not in lines
    assert "5.2" not in lines
    names = " ".join(r["name"] for r in rows)
    assert "миндалин" not in names
    assert any(r["name"] == "Всего операций" and r["total"] == 3 for r in rows)
