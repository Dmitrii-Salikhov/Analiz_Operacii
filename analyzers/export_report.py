# analyzers/export_report.py
"""Экспорт месяца в копии шаблона сводной (тот же вид и формулы)."""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import List, Optional

import pandas as pd

from analyzers.summary_writer import SummaryWriter


def export_month_like_summary(
    template_path: str,
    output_path: str,
    ops: pd.DataFrame,
    summary_cfg: dict,
    department: str,
    categories: List[dict],
    pension_age: int = 60,
    month: Optional[int] = None,
    year: Optional[int] = None,
) -> dict:
    """
    Копирует шаблон сводной и записывает в него операции (обычно один месяц).
    Стиль, формулы ИТОГ / план / форма 4001 — как в рабочем файле.
    """
    src = Path(template_path)
    out = Path(output_path)
    if not src.exists():
        raise FileNotFoundError(src)
    shutil.copy2(src, out)

    df = ops.copy()
    if month is not None and not df.empty and "Дата" in df.columns:
        dates = pd.to_datetime(df["Дата"], errors="coerce")
        mask = dates.dt.month == int(month)
        if year is not None:
            mask &= dates.dt.year == int(year)
        df = df.loc[mask].copy()

    if df.empty:
        d_min = d_max = None
    else:
        dates = pd.to_datetime(df["Дата"], errors="coerce")
        d_min, d_max = dates.min(), dates.max()

    writer = SummaryWriter(
        str(out),
        summary_cfg,
        department=department,
        categories=categories,
        pension_age=pension_age,
    )
    return writer.write(
        df,
        output_path=str(out),
        overwrite_from=d_min,
        overwrite_to=d_max,
        backup=False,
        write_weeks=True,
        write_form=True,
    )
