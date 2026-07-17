# analyzers/io_utils.py
"""Чтение отчётов и накопление операций с перезаписью пересекающихся недель."""
from __future__ import annotations

import os
from typing import Iterable, List, Optional, Set, Tuple

import pandas as pd


def smart_read_excel(file_path: str) -> pd.DataFrame:
    """Читает Excel, подбирая строку заголовка (0–2) с минимумом Unnamed."""
    best_df = None
    best_unnamed = float("inf")
    for header_row in (0, 1, 2):
        try:
            df = pd.read_excel(file_path, header=header_row)
            unnamed_count = sum("Unnamed" in str(col) for col in df.columns)
            if unnamed_count < best_unnamed:
                best_unnamed = unnamed_count
                best_df = df
            if unnamed_count == 0:
                break
        except Exception:
            continue
    if best_df is None:
        raise ValueError("Не удалось прочитать файл с заголовками строк 0–2")
    return best_df.dropna(how="all").dropna(axis=1, how="all")


def read_table(file_path: str) -> pd.DataFrame:
    if file_path.lower().endswith(".csv"):
        return pd.read_csv(file_path, encoding="utf-8")
    return smart_read_excel(file_path)


def find_column(df: pd.DataFrame, keywords: Iterable[str]) -> Optional[str]:
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if all(kw.lower() in col_lower for kw in keywords):
            return col
    return None


def week_start(dt: pd.Timestamp) -> pd.Timestamp:
    """Понедельник недели операции (для ключа перезаписи)."""
    d = pd.Timestamp(dt).normalize()
    return d - pd.Timedelta(days=int(d.weekday()))


def ops_dedupe_key(row: pd.Series) -> Tuple:
    return (
        str(row.get("КВС", "")),
        pd.Timestamp(row["Дата"]).normalize() if pd.notna(row.get("Дата")) else None,
        str(row.get("Код", "")),
        str(row.get("Категория", "")),
    )


class OperationsStore:
    """Накопитель операций: при пересечении периода — данные нового отчёта вытесняют старые."""

    def __init__(self):
        self.ops: pd.DataFrame = pd.DataFrame()
        self.sources: List[str] = []
        # basename → {date_from, date_to, count}
        self.source_meta: dict = {}

    def clear(self):
        self.ops = pd.DataFrame()
        self.sources = []
        self.source_meta = {}

    def covered_weeks(self, ops: pd.DataFrame) -> Set[pd.Timestamp]:
        if ops.empty or "Дата" not in ops.columns:
            return set()
        return set(ops["Дата"].dropna().map(week_start))

    def date_span(self, ops: pd.DataFrame):
        if ops is None or ops.empty or "Дата" not in ops.columns:
            return None, None
        dates = pd.to_datetime(ops["Дата"], errors="coerce").dropna()
        if dates.empty:
            return None, None
        return dates.min().normalize(), dates.max().normalize()

    def refresh_source_meta(self):
        """Пересчитать периоды по фактическим данным накопителя."""
        meta = {}
        if self.ops.empty or "_source" not in self.ops.columns:
            self.source_meta = meta
            return
        for src, subset in self.ops.groupby(self.ops["_source"].astype(str)):
            d0, d1 = self.date_span(subset)
            meta[str(src)] = {
                "date_from": d0,
                "date_to": d1,
                "count": len(subset),
            }
        self.source_meta = meta
        # порядок sources сохраняем, убираем пустые
        self.sources = [s for s in self.sources if s in meta] + [
            s for s in meta if s not in self.sources
        ]

    def add(self, new_ops: pd.DataFrame, source_path: str) -> dict:
        """
        Добавляет операции. Все записи store, попадающие в диапазон дат нового
        отчёта [min, max], удаляются и заменяются данными из new_ops.
        """
        if new_ops is None or new_ops.empty:
            return {"added": 0, "removed": 0, "weeks": 0, "total": 0}

        new_ops = new_ops.copy()
        new_ops["_source"] = os.path.basename(source_path)
        d_min, d_max = self.date_span(new_ops)
        weeks = self.covered_weeks(new_ops)
        removed = 0

        if not self.ops.empty and d_min is not None:
            old_dates = pd.to_datetime(self.ops["Дата"], errors="coerce")
            mask_drop = (old_dates >= d_min) & (old_dates <= d_max)
            removed = int(mask_drop.sum())
            self.ops = self.ops.loc[~mask_drop].copy()

        combined = pd.concat([self.ops, new_ops], ignore_index=True) if not self.ops.empty else new_ops
        keys = combined.apply(ops_dedupe_key, axis=1)
        combined = combined.loc[~keys.duplicated(keep="last")].copy()
        self.ops = combined.reset_index(drop=True)

        base = os.path.basename(source_path)
        if base not in self.sources:
            self.sources.append(base)
        self.refresh_source_meta()

        return {
            "added": len(new_ops),
            "removed": removed,
            "weeks": len(weeks),
            "date_from": d_min,
            "date_to": d_max,
            "total": len(self.ops),
        }
