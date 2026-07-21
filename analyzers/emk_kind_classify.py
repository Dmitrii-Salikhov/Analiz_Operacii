# analyzers/emk_kind_classify.py
"""Классификация категорий план/экстренно по данным ЭМК; спорные — на ручную сверку."""
from __future__ import annotations

from collections import Counter, defaultdict
from typing import Dict, List, Optional, Tuple

import pandas as pd

from analyzers.emk_loader import normalize_hosp_type


def category_emk_counts(ops_df: pd.DataFrame) -> Dict[str, Counter]:
    """Категория → {план, экстр, нет} по полю Тип_ЭМК."""
    result: Dict[str, Counter] = defaultdict(Counter)
    if ops_df is None or ops_df.empty:
        return {}
    for _, row in ops_df.iterrows():
        cat = str(row.get("Категория") or "").strip()
        if not cat or cat == "Не классифицировано":
            continue
        kind = normalize_hosp_type(row.get("Тип_ЭМК"))
        if kind == "плановая":
            result[cat]["план"] += 1
        elif kind == "экстренная":
            result[cat]["экстр"] += 1
        else:
            result[cat]["нет"] += 1
    return dict(result)


def classify_categories_by_emk(
    ops_df: pd.DataFrame,
    *,
    category_names: Optional[List[str]] = None,
    min_linked: int = 1,
) -> dict:
    """
    По операциям с Тип_ЭМК:
    - plan: все связанные эпизоды — плановые
    - emergency: все связанные — экстренные
    - disputed: есть и план, и экстр (на ручную сверку)
    - no_emk: нет ни одного связанного эпизода

    Возвращает:
      plan, emergency, disputed (list of dict), no_emk,
      plan_categories, emergency_categories (для summary_cfg)
    """
    counts = category_emk_counts(ops_df)
    names = list(category_names) if category_names is not None else sorted(counts.keys())

    plan: List[str] = []
    emergency: List[str] = []
    disputed: List[dict] = []
    no_emk: List[str] = []

    for cat in names:
        c = counts.get(cat) or Counter()
        p = int(c.get("план", 0))
        e = int(c.get("экстр", 0))
        n = int(c.get("нет", 0))
        linked = p + e
        if linked < min_linked:
            no_emk.append(cat)
            plan.append(cat)  # по умолчанию план, пока нет ЭМК
            continue
        if e == 0:
            plan.append(cat)
        elif p == 0:
            emergency.append(cat)
        else:
            disputed.append(
                {
                    "Категория": cat,
                    "План_ЭМК": p,
                    "Экстр_ЭМК": e,
                    "Без_типа": n,
                    "Доля_экстр_%": round(100.0 * e / linked),
                    "Рекомендация": "экстренная" if e > p else ("плановая" if p > e else "уточнить"),
                }
            )
            # до ручного решения — в ту сторону, где больше; при равенстве — план
            if e > p:
                emergency.append(cat)
            else:
                plan.append(cat)

    return {
        "plan": plan,
        "emergency": emergency,
        "disputed": disputed,
        "no_emk": no_emk,
        "plan_categories": list(plan),
        "emergency_categories": list(emergency),
        "counts": {k: dict(v) for k, v in counts.items()},
    }


def apply_kind_to_summary_cfg(summary_cfg: dict, classification: dict) -> dict:
    """Обновляет plan_categories / emergency_categories в summary_cfg."""
    out = dict(summary_cfg)
    out["plan_categories"] = list(classification.get("plan_categories") or [])
    out["emergency_categories"] = list(classification.get("emergency_categories") or [])
    return out


def disputed_to_dataframe(classification: dict) -> pd.DataFrame:
    rows = classification.get("disputed") or []
    if not rows:
        return pd.DataFrame(
            columns=["Категория", "План_ЭМК", "Экстр_ЭМК", "Без_типа", "Доля_экстр_%", "Рекомендация"]
        )
    return pd.DataFrame(rows)


def format_kind_report(classification: dict, *, limit_disputed: int = 40) -> str:
    lines = [
        "Классификация план/экстренно по ЭМК",
        f"  Плановые: {len(classification.get('plan') or [])}",
        f"  Экстренные: {len(classification.get('emergency') or [])}",
        f"  Без ЭМК (оставлены план): {len(classification.get('no_emk') or [])}",
        f"  Спорные (план и экстр): {len(classification.get('disputed') or [])}",
    ]
    disputed = classification.get("disputed") or []
    if disputed:
        lines.append("")
        lines.append("Спорные — нужна ручная сверка:")
        for d in disputed[:limit_disputed]:
            lines.append(
                f"  • {d['Категория']}: план={d['План_ЭМК']}, экстр={d['Экстр_ЭМК']} "
                f"({d['Доля_экстр_%']}% экстр) → рек.: {d['Рекомендация']}"
            )
        if len(disputed) > limit_disputed:
            lines.append(f"  … и ещё {len(disputed) - limit_disputed}")
    return "\n".join(lines)
