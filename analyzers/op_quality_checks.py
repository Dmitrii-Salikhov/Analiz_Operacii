# analyzers/op_quality_checks.py
"""Проверки качества опержурнала: длительные операции и пустой опер. стол."""
from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

DEFAULT_LONG_OP_HOURS = 4.0


def _fmt_kvs(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none"):
        return ""
    # Excel иногда даёт 26.22222 как float — оставляем строку как в журнале
    return s


def _fmt_hours(val: Any) -> str:
    try:
        return f"{float(val):.2f}"
    except (TypeError, ValueError):
        return ""


def _base_row(r: pd.Series, *, reason: str, hours: Any = None) -> dict:
    return {
        "КВС": _fmt_kvs(r.get("КВС")),
        "Пациент": str(r.get("Пациент") or "").strip(),
        "Хирург": str(r.get("Хирург") or "").strip(),
        "Услуга": str(r.get("Услуга") or "").strip(),
        "Длительность_ч": hours if hours is not None else r.get("Длительность_ч"),
        "Длительность": _fmt_hours(hours if hours is not None else r.get("Длительность_ч")),
        "Причина": reason,
        "Опер.стол": str(r.get("Опер.стол") or "").strip(),
        "Дата": r.get("Дата"),
    }


def find_long_operations(
    ops: pd.DataFrame,
    *,
    max_hours: float = DEFAULT_LONG_OP_HOURS,
) -> List[dict]:
    """Операции с длительностью строго больше max_hours."""
    if ops is None or getattr(ops, "empty", True):
        return []
    if "Длительность_ч" not in ops.columns:
        return []
    thr = float(max_hours)
    out: List[dict] = []
    seen = set()
    for _, r in ops.iterrows():
        hours = r.get("Длительность_ч")
        if hours is None or (isinstance(hours, float) and pd.isna(hours)):
            continue
        try:
            h = float(hours)
        except (TypeError, ValueError):
            continue
        if h <= thr:
            continue
        key = (
            _fmt_kvs(r.get("КВС")),
            str(r.get("Услуга") or "")[:80],
            _fmt_hours(h),
            str(r.get("Начало") or ""),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(
            _base_row(
                r,
                reason=f"длительность > {thr:g} ч",
                hours=h,
            )
        )
    return out


def find_missing_or_table(ops: pd.DataFrame) -> List[dict]:
    """Услуга есть, а «Опер.стол» пустой / только пробелы."""
    if ops is None or getattr(ops, "empty", True):
        return []
    out: List[dict] = []
    seen = set()
    for _, r in ops.iterrows():
        service = str(r.get("Услуга") or "").strip()
        if not service or service.lower() in ("nan", "none"):
            continue
        table = str(r.get("Опер.стол") or "").strip()
        if table and table.lower() not in ("nan", "none"):
            continue
        key = (_fmt_kvs(r.get("КВС")), service[:80], str(r.get("Дата") or ""))
        if key in seen:
            continue
        seen.add(key)
        out.append(_base_row(r, reason="не занесена на опер.стол"))
    return out


def long_op_hours_from_config(config: Optional[dict]) -> float:
    thr = (config or {}).get("thresholds") or {}
    try:
        return float(thr.get("long_op_hours", DEFAULT_LONG_OP_HOURS))
    except (TypeError, ValueError):
        return DEFAULT_LONG_OP_HOURS
