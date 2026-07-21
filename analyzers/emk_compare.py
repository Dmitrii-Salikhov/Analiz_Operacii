# analyzers/emk_compare.py
"""Сверка план/экстренно: шаблон (по категориям) vs тип госпитализации из ЭМК."""
from __future__ import annotations

from typing import Dict, List, Optional

import pandas as pd

from analyzers.emk_loader import normalize_hosp_type


def template_kind(category: str, summary_cfg: dict) -> str:
    if category in summary_cfg.get("emergency_categories", []):
        return "экстренная"
    if category in summary_cfg.get("plan_categories", []):
        return "плановая"
    return "неизвестно"


def emk_kind_from_row(row: pd.Series) -> str:
    return normalize_hosp_type(row.get("Тип_ЭМК"))


def compare_plan_emergency(
    ops_df: pd.DataFrame,
    summary_cfg: dict,
    *,
    department: Optional[str] = None,
) -> Dict:
    """
    Сравнивает классификацию шаблона с Типом_ЭМК у каждой операции.
    department — для подписи в отчёте (фильтрация ops_df снаружи).
    """
    if ops_df is None or ops_df.empty:
        return _empty_result(department)

    mismatches: List[dict] = []
    t_plan = t_emerg = t_unknown = 0
    e_plan = e_emerg = e_unk = 0
    match = compared = linked = 0

    for _, row in ops_df.iterrows():
        cat = str(row.get("Категория", ""))
        t_kind = template_kind(cat, summary_cfg)
        if t_kind == "плановая":
            t_plan += 1
        elif t_kind == "экстренная":
            t_emerg += 1
        else:
            t_unknown += 1

        e_kind = emk_kind_from_row(row)
        if e_kind == "плановая":
            e_plan += 1
            linked += 1
        elif e_kind == "экстренная":
            e_emerg += 1
            linked += 1
        else:
            e_unk += 1

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

    all_plan_template = t_emerg == 0 and t_plan > 0 and not summary_cfg.get("emergency_categories")
    return {
        "department": department or "",
        "template_plan": t_plan,
        "template_emerg": t_emerg,
        "template_unknown": t_unknown,
        "emk_plan": e_plan,
        "emk_emerg": e_emerg,
        "emk_unknown": e_unk,
        "emk_linked": linked,
        "mismatches": mismatches,
        "match_count": match,
        "compared": compared,
        "all_plan_template": all_plan_template,
        "total_ops": len(ops_df),
    }


def _empty_result(department: Optional[str] = None) -> Dict:
    return {
        "department": department or "",
        "template_plan": 0,
        "template_emerg": 0,
        "template_unknown": 0,
        "emk_plan": 0,
        "emk_emerg": 0,
        "emk_unknown": 0,
        "emk_linked": 0,
        "mismatches": [],
        "match_count": 0,
        "compared": 0,
        "all_plan_template": False,
        "total_ops": 0,
    }


def format_mismatch_report(result: Dict, limit: int = 50) -> str:
    dept = str(result.get("department") or "").strip()
    lines = [
        "Сверка план/экстренно (шаблон vs ЭМК)",
    ]
    if dept:
        lines.append(f"  Отделение: {dept}")
    lines.extend(
        [
            f"  Операций всего: {result.get('total_ops', 0)}",
            f"  Связано с ЭМК: {result.get('emk_linked', 0)} (без типа: {result.get('emk_unknown', 0)})",
            f"  По шаблону: план={result['template_plan']}, экстр={result['template_emerg']}",
            f"  По ЭМК:     план={result['emk_plan']}, экстр={result['emk_emerg']}",
            f"  Совпало: {result['match_count']} из {result['compared']}",
            f"  Расхождений: {len(result['mismatches'])}",
        ]
    )
    if result.get("all_plan_template") and result.get("emk_emerg", 0) > 0:
        lines.append(
            "  Примечание: в шаблоне все операции = план; расхождения с экстренной ЭМК ожидаемы "
            "до настройки plan/emergency по категориям."
        )
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
