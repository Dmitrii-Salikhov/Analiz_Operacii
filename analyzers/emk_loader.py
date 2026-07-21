# analyzers/emk_loader.py
"""Чтение отчёта «Заполнение ЭМК в стационаре (суммарно)»."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd

from analyzers.io_utils import find_column, smart_read_excel

# Маркеры формата «Отчет по заполнению ЭМК в стационаре»
_EMK_MARKERS = ("номер квс", "тип госпитализации", "дата выписки из стационара")


def _row_looks_like_header(row: pd.Series) -> bool:
    text = " ".join(str(v).lower() for v in row.values if pd.notna(v))
    return sum(m in text for m in _EMK_MARKERS) >= 2


def detect_emk_header_row(path: str | Path, max_scan: int = 5) -> int:
    """Номер строки заголовка (0-based) для отчёта ЭМК."""
    path = Path(path)
    preview = pd.read_excel(path, header=None, nrows=max_scan)
    for i in range(len(preview)):
        if _row_looks_like_header(preview.iloc[i]):
            return int(i)
    return 2 if len(preview) > 2 else 0


def normalize_emk_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Убирает лишние пробелы в названиях колонок."""
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def read_emk_stationary_report(path: str | Path) -> pd.DataFrame:
    """
    Читает xlsx/csv отчёт по заполнению ЭМК в стационаре.
    Заголовок обычно на 3-й строке Excel (header=2).
    """
    path = Path(path)
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, encoding="utf-8")
        return normalize_emk_columns(df.dropna(how="all"))

    header_row = detect_emk_header_row(path)
    df = pd.read_excel(path, header=header_row)
    df = normalize_emk_columns(df.dropna(how="all").dropna(axis=1, how="all"))

    required = find_column(df, ["номер", "квс"])
    if not required:
        # fallback — универсальное чтение
        return normalize_emk_columns(smart_read_excel(str(path)))
    return df


def filter_emk_by_department(df: pd.DataFrame, department: str) -> pd.DataFrame:
    """Оставляет строки ЭМК, где «Отделение» содержит имя отделения."""
    if df is None or df.empty or not department:
        return df
    col = find_column(df, ["отделение"]) or "Отделение"
    if col not in df.columns:
        return df
    mask = df[col].astype(str).str.contains(str(department).strip(), na=False, regex=False)
    return df.loc[mask].copy()


def emk_department_stats(df: pd.DataFrame) -> dict:
    """Краткая сводка по отделениям в загруженном ЭМК."""
    if df is None or df.empty:
        return {}
    col = find_column(df, ["отделение"]) or "Отделение"
    if col not in df.columns:
        return {"_all": len(df)}
    vc = df[col].astype(str).value_counts()
    return {str(k): int(v) for k, v in vc.items()}


def normalize_hosp_type(raw: Optional[str]) -> str:
    s = str(raw or "").strip().lower()
    if s.startswith("экстр"):
        return "экстренная"
    if s.startswith("план"):
        return "плановая"
    return ""
