# analyzers/emk_compare.py
"""Сверка план/экстренно: шаблон (по категориям) vs тип госпитализации из ЭМК."""
from __future__ import annotations

from typing import Dict, List

import pandas as pd


def template_kind(category: str, summary_cfg: dict) -> str:
    if category in summary_cfg.get("emergency_categories", []):
        return "экстренная"
    if category in summary_cfg.get("plan_categories", []):
        return "плановая"
    return "неизвестно"


def compare_plan_emergency(ops_df: pd.DataFrame, summary_cfg: dict) -> Dict:
    """
    Сравнивает классификацию шаблона с Типом_ЭМК у каждой операции.
    Возвращает сводку и список расхождений.
    """
    if ops_df is None or ops_df.empty:
        return {
            "template_plan": 0,
            "template_emerg": 0,
            "emk_plan": 0,
            "emk_emerg": 0,
            "emk_unknown": 0,
            "mismatches": [],
            "match_count": 0,
            "compared": 0,
        }

    mismatches: List[dict] = []
    t_plan = t_emerg = 0
    e_plan = e_emerg = e_unk = 0
    match = compared = 0

    for _, row in ops_df.iterrows():
        cat = str(row.get("Категория", ""))
        t_kind = template_kind(cat, summary_cfg)
        if t_kind == "плановая":
            t_plan += 1
        elif t_kind == "экстренная":
            t_emerg += 1

        emk = str(row.get("Тип_ЭМК", "") or "").strip().lower()
        if emk.startswith("план"):
            e_plan += 1
            e_kind = "плановая"
        elif emk.startswith("экстр"):
            e_emerg += 1
            e_kind = "экстренная"
        else:
            e_unk += 1
            e_kind = ""

        if not e_kind:
            continue
        compared += 1
        if e_kind == t_kind:
            match += 1
        else:
            mismatches.append(
                {
                    "КВС": row.get("КВС"),
                    "Дата": row.get("Дата"),
                    "Категория": cat,
                    "Код": row.get("Код"),
                    "Услуга": row.get("Услуга"),
                    "Диагноз": row.get("Диагноз"),
                    "Шаблон": t_kind,
                    "ЭМК": e_kind,
                }
            )

    return {
        "template_plan": t_plan,
        "template_emerg": t_emerg,
        "emk_plan": e_plan,
        "emk_emerg": e_emerg,
        "emk_unknown": e_unk,
        "mismatches": mismatches,
        "match_count": match,
        "compared": compared,
    }


def format_mismatch_report(result: Dict, limit: int = 50) -> str:
    lines = [
        "Сверка план/экстренно (шаблон vs ЭМК)",
        f"  По шаблону: план={result['template_plan']}, экстр={result['template_emerg']}",
        f"  По ЭМК:     план={result['emk_plan']}, экстр={result['emk_emerg']}, без типа={result['emk_unknown']}",
        f"  Совпало: {result['match_count']} из {result['compared']}",
        f"  Расхождений: {len(result['mismatches'])}",
    ]
    if result["mismatches"]:
        lines.append("")
        lines.append("Детали расхождений:")
        for m in result["mismatches"][:limit]:
            dt = m["Дата"]
            dt_s = dt.strftime("%d.%m.%Y") if hasattr(dt, "strftime") else str(dt)
            lines.append(
                f"  • КВС {m['КВС']} | {dt_s} | {m['Категория']} ({m['Код']})"
                f" | шаблон={m['Шаблон']}, ЭМК={m['ЭМК']}"
            )
            if m.get("Диагноз"):
                diag = str(m["Диагноз"])
                if len(diag) > 100:
                    diag = diag[:100] + "…"
                lines.append(f"      диагноз: {diag}")
            if m.get("Услуга"):
                svc = str(m["Услуга"])
                if len(svc) > 100:
                    svc = svc[:100] + "…"
                lines.append(f"      услуга: {svc}")
        if len(result["mismatches"]) > limit:
            lines.append(f"  … и ещё {len(result['mismatches']) - limit}")
    return "\n".join(lines)
