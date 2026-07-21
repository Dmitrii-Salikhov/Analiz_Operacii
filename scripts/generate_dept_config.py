#!/usr/bin/env python3
"""Генерация summaries.* / surgery_categories_by_dept.* + план/экстр по ЭМК."""
from __future__ import annotations

import sys
from copy import deepcopy
from pathlib import Path

import yaml

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.dept_config import DEPT_REPORT_SOURCES, ensure_multi_dept_config
from analyzers.dept_inventory import (
    build_categories_from_codes,
    build_inventory_table,
    build_summary_cfg_draft,
)
from analyzers.dept_template import create_from_summary_cfg
from analyzers.emk_kind_classify import (
    apply_kind_to_summary_cfg,
    classify_categories_by_emk,
    disputed_to_dataframe,
    format_kind_report,
)
from analyzers.emk_loader import read_emk_stationary_report
from analyzers.io_utils import smart_read_excel
from analyzers.surgery import SurgeryAnalyzer

REPORTS = APP / "Отчеты других отделений"
CONFIG = APP / "config.yaml"
EMK_CANDIDATES = [
    APP / "Отчет по заполнению ЭМК в стационаре(суммарно) (1).xlsx",
    APP / "Отчет по заполнению ЭМК в стационаре(суммарно).xlsx",
]
DEPT_KEYS = ("surg1", "surg2", "pedsurg", "traum")


def _find_emk() -> Path | None:
    for p in EMK_CANDIDATES:
        if p.exists():
            return p
    return None


def main() -> None:
    with CONFIG.open("r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    ensure_multi_dept_config(config)

    summaries = config.setdefault("summaries", {})
    by_dept = config.setdefault("surgery_categories_by_dept", {})

    if config.get("summary"):
        summaries["lor"] = deepcopy(config["summary"])
    if config.get("surgery_categories"):
        by_dept["lor"] = deepcopy(config["surgery_categories"])

    emk_path = _find_emk()
    emk_df = read_emk_stationary_report(emk_path) if emk_path else None
    if emk_path:
        print(f"ЭМК: {emk_path.name} ({len(emk_df)} стр.)")
    else:
        print("ЭМК: не найден — все категории останутся плановыми")

    disputed_all = []

    for key in DEPT_KEYS:
        meta = DEPT_REPORT_SOURCES[key]
        table = build_inventory_table(REPORTS / meta["report_file"], meta["department"])
        cats, _rows, smeta = build_categories_from_codes(table, group=meta.get("label", "операции"))
        summary_cfg = build_summary_cfg_draft(key, cats, smeta, year=2026)

        if emk_df is not None:
            df = smart_read_excel(str(REPORTS / meta["report_file"]))
            ops = SurgeryAnalyzer(df, meta["department"], cats, emk_df=emk_df).extract_operations()
            names = [c["category"] for c in cats]
            kind = classify_categories_by_emk(ops, category_names=names)
            summary_cfg = apply_kind_to_summary_cfg(summary_cfg, kind)
            print(format_kind_report(kind))
            for d in kind.get("disputed") or []:
                disputed_all.append({"Отделение": meta["department"], "summary_key": key, **d})
            print(
                f"{key}: {len(cats)} кат. | план={len(kind['plan'])} экстр={len(kind['emergency'])} "
                f"спор={len(kind['disputed'])} без_эмк={len(kind['no_emk'])}"
            )
        else:
            print(f"{key}: {len(cats)} categories (все plan)")

        summaries[key] = summary_cfg
        by_dept[key] = cats

        # пересоздать Excel-сводную
        out = APP / summary_cfg["default_path"]
        create_from_summary_cfg(out, summary_cfg, meta["department"])
        print(f"  → {out.name}")

    config["summaries"] = summaries
    config["surgery_categories_by_dept"] = by_dept
    config["summary"] = summaries["lor"]
    config["surgery_categories"] = by_dept["lor"]

    text = yaml.safe_dump(config, allow_unicode=True, default_flow_style=False, sort_keys=False, width=120)
    CONFIG.write_text(text, encoding="utf-8")
    print(f"Updated {CONFIG}")

    if disputed_all:
        out_disp = APP / "спорные_план_экстр_ЭМК.xlsx"
        disputed_to_dataframe({"disputed": disputed_all}).to_excel(out_disp, index=False)
        print(f"Спорные для ручной сверки: {out_disp} ({len(disputed_all)})")
    else:
        print("Спорных план/экстр нет")


if __name__ == "__main__":
    main()
