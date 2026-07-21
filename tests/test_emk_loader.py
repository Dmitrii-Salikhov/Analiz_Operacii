# tests/test_emk_loader.py
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.emk_compare import compare_plan_emergency, format_mismatch_report
from analyzers.emk_loader import detect_emk_header_row, read_emk_stationary_report
from analyzers.io_utils import smart_read_excel
from analyzers.surgery import SurgeryAnalyzer, resolve_emk_hosp_type
from analyzers.dept_config import get_surgery_categories, get_summary_cfg
import yaml


def test_detect_emk_header_row():
    path = APP / "Отчет по заполнению ЭМК в стационаре(суммарно) (1).xlsx"
    if not path.exists():
        return
    assert detect_emk_header_row(path) == 2


def test_read_emk_stationary_report():
    path = APP / "Отчет по заполнению ЭМК в стационаре(суммарно) (1).xlsx"
    if not path.exists():
        return
    df = read_emk_stationary_report(path)
    assert "Номер КВС" in df.columns
    assert "Тип госпитализации" in df.columns
    assert len(df) > 100


def test_emk_link_surg1():
    emk_path = APP / "Отчет по заполнению ЭМК в стационаре(суммарно) (1).xlsx"
    report = APP / "Отчеты других отделений/1 хир.xlsx"
    cfg_path = APP / "config.yaml"
    if not all(p.exists() for p in (emk_path, report, cfg_path)):
        return
    with cfg_path.open(encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    emk = read_emk_stationary_report(emk_path)
    df = smart_read_excel(str(report))
    dept = "1 Хирургическое отделение"
    cats = get_surgery_categories(cfg, summary_key="surg1")
    ops = SurgeryAnalyzer(df, dept, cats, emk_df=emk).extract_operations()
    linked = ops["Тип_ЭМК"].astype(str).str.strip().ne("").sum()
    assert linked >= 50  # улучшенная привязка по КВС
    result = compare_plan_emergency(
        ops, get_summary_cfg(cfg, summary_key="surg1"), department=dept
    )
    assert result["compared"] == linked
    text = format_mismatch_report(result, limit=3)
    assert "Сверка план/экстренно" in text


def test_resolve_emk_hosp_type_single_episode():
    import pandas as pd

    episodes = [
        {
            "admission": pd.Timestamp("2026-06-10"),
            "discharge": pd.Timestamp("2026-06-11"),
            "type": "экстренная",
            "diagnosis": "X",
        }
    ]
    # операция накануне расчётного поступления — grace 1 день
    assert resolve_emk_hosp_type(episodes, pd.Timestamp("2026-06-09")) == "экстренная"
    assert resolve_emk_hosp_type(episodes, pd.Timestamp("2026-06-10")) == "экстренная"
