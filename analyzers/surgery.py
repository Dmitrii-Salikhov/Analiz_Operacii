# analyzers/surgery.py
from __future__ import annotations

import re
from datetime import timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from analyzers.io_utils import find_column
from analyzers.ksg_catalog import get_catalog


class SurgeryAnalyzer:
    def __init__(self, df: pd.DataFrame, department: str, categories_config: List[dict], emk_df=None):
        mask = (
            df["Отделение госпитализации"].astype(str).str.contains(department, na=False)
            | df["Оперблок"].astype(str).str.contains(department, na=False)
        )
        self.df = df[mask].copy()
        self.department = department
        self.categories = categories_config
        self.emk_df = emk_df
        self.ksg = get_catalog()

        date_start_col = find_column(self.df, ["дата", "начала"])
        if not date_start_col:
            raise KeyError("Не найдена колонка с датой начала операции")
        self.df["Дата операции"] = pd.to_datetime(self.df[date_start_col], dayfirst=True, errors="coerce")

        birth_col = find_column(self.df, ["дата", "рождения"])
        if not birth_col:
            raise KeyError("Не найдена колонка с датой рождения пациента")
        self.df["Дата рождения"] = pd.to_datetime(self.df[birth_col], dayfirst=True, errors="coerce")

        self.df["Возраст"] = self.df.apply(
            lambda row: int((row["Дата операции"] - row["Дата рождения"]).days // 365)
            if pd.notna(row["Дата операции"]) and pd.notna(row["Дата рождения"])
            else None,
            axis=1,
        )

        self._code_index = self._build_code_index()
        self._emk_hosp_index = self._build_emk_hosp_index()

    def _build_code_index(self) -> Dict[str, dict]:
        index = {}
        for cat in self.categories:
            for code in cat.get("codes") or []:
                index[code] = cat
        return index

    def _build_emk_hosp_index(self) -> Dict[Any, List[dict]]:
        """КВС → список госпитализаций из ЭМК (для миринготомии и сверки)."""
        if self.emk_df is None or self.emk_df.empty:
            return {}
        emk = self.emk_df
        if "Номер КВС" not in emk.columns:
            return {}
        result: Dict[Any, List[dict]] = {}
        for _, row in emk.iterrows():
            kvs = str(row["Номер КВС"]).strip()
            if not kvs:
                continue
            try:
                discharge = pd.to_datetime(row.get("Дата выписки из стационара"), dayfirst=True, errors="coerce")
                bed_days = int(row.get("Всего дней проведено в стационаре (от поступления до исхода в днях)", 1) or 1)
                if bed_days < 1:
                    bed_days = 1
                if pd.isna(discharge):
                    continue
                admission = discharge - timedelta(days=bed_days) + timedelta(days=1)
                typ = str(row.get("Тип госпитализации", "")).strip().lower()
                diag = str(row.get("Основной диагноз", "") or "")
                result.setdefault(kvs, []).append(
                    {
                        "admission": admission,
                        "discharge": discharge,
                        "type": typ,
                        "diagnosis": diag,
                    }
                )
            except Exception:
                continue
        return result

    def hosp_type_for(self, kvs, date_op) -> Optional[str]:
        episodes = self._emk_hosp_index.get(str(kvs).strip(), [])
        if not episodes or pd.isna(date_op):
            return None
        for ep in episodes:
            if ep["admission"] <= date_op <= ep["discharge"]:
                return ep["type"]
        return None

    def diagnosis_for(self, kvs, date_op) -> str:
        episodes = self._emk_hosp_index.get(str(kvs).strip(), [])
        if not episodes or pd.isna(date_op):
            return ""
        for ep in episodes:
            if ep["admission"] <= date_op <= ep["discharge"]:
                return ep.get("diagnosis", "")
        return ""

    def extract_operations(self) -> pd.DataFrame:
        ops = []
        for _, row in self.df.iterrows():
            if pd.isna(row.get("Дата операции")):
                continue
            text = str(row.get("Услуга", "") or "")
            codes = re.findall(r"A\d{2}\.\d{2}\.\d{3}(?:\.\d{3})?", text)
            team = str(row.get("Операционная бригада", "") or "")
            surgeon = self._extract_surgeon(team)
            table = str(row["Опер.стол"]).strip() if pd.notna(row.get("Опер.стол")) else ""
            kvs = row.get("№ истории")
            date_op = row["Дата операции"]
            hosp_type = self.hosp_type_for(kvs, date_op)
            diagnosis = self.diagnosis_for(kvs, date_op)

            if codes:
                for code in codes:
                    cat_info = self._classify(code, text, hosp_type, companion_codes=codes)
                    ops.append(self._op_dict(row, code, cat_info, surgeon, table, hosp_type, diagnosis, text))
            else:
                cat_info = self._classify_by_name(text, hosp_type)
                if cat_info is None:
                    continue
                ops.append(self._op_dict(row, "", cat_info, surgeon, table, hosp_type, diagnosis, text))

        result = pd.DataFrame(ops)
        return self._force_myringotomy_plan_with_adenotomy(result)

    def _op_dict(self, row, code, cat_info, surgeon, table, hosp_type, diagnosis, service_text):
        if cat_info is None:
            cat_name, group, line, hist = "Не классифицировано", "прочее", "", False
        else:
            cat_name, group, line, hist = cat_info
        ksg_name = self.ksg.name_for(code) if code else ""
        ksg_groups = self.ksg.ksg_for(code) if code else ""
        ksg_hint = self.ksg.hint_for(code) if code and cat_name == "Не классифицировано" else ""
        return {
            "Дата": row["Дата операции"],
            "КВС": row.get("№ истории"),
            "Дата рождения": row.get("Дата рождения"),
            "Возраст": row.get("Возраст"),
            "Отделение": row.get("Отделение госпитализации"),
            "Опер.стол": table,
            "Хирург": surgeon,
            "Код": code,
            "Услуга": service_text,
            "Категория": cat_name,
            "Группа": group,
            "Строка_4001": line,
            "Гистология": hist,
            "Тип_ЭМК": hosp_type or "",
            "Диагноз": diagnosis,
            "КСГ_название": ksg_name,
            "КСГ": ksg_groups,
            "КСГ_подсказка": ksg_hint,
        }

    # Код аденотомии: миринготомия в одной операции / в тот же день → всегда плановая
    ADENOTOMY_CODES = {"A16.08.002.001"}

    def _classify(self, code: str, service_text: str, hosp_type: Optional[str], companion_codes=None):
        cat = self._code_index.get(code)
        if cat is None:
            return self._classify_by_name(service_text, hosp_type, companion_codes=companion_codes)
        return self._resolve_category(cat, hosp_type, companion_codes=companion_codes, service_text=service_text)

    def _classify_by_name(self, service_text: str, hosp_type: Optional[str], companion_codes=None):
        text = (service_text or "").lower()
        best = None
        best_score = 0
        for cat in self.categories:
            kws = cat.get("name_keywords") or []
            if not kws:
                continue
            score = sum(1 for kw in kws if kw.lower() in text)
            if score > best_score:
                best_score = score
                best = cat
        if best is None or best_score == 0:
            return None
        return self._resolve_category(best, hosp_type, companion_codes=companion_codes, service_text=service_text)

    def _has_adenotomy(self, companion_codes=None, service_text: str = "") -> bool:
        codes = set(companion_codes or [])
        if codes & self.ADENOTOMY_CODES:
            return True
        text = (service_text or "").lower()
        return "аденоид" in text or "аденотоми" in text

    def _resolve_category(
        self,
        cat: dict,
        hosp_type: Optional[str],
        companion_codes=None,
        service_text: str = "",
    ) -> Tuple[str, str, str, bool]:
        name = cat["category"]
        emerg_alt = cat.get("emergency_category")
        # Миринготомия вместе с аденотомией — всегда плановая
        if emerg_alt and self._has_adenotomy(companion_codes, service_text):
            name = cat["category"]  # «Миринготомия план»
        elif emerg_alt and hosp_type == "экстренная":
            name = emerg_alt
        return name, cat.get("group", ""), cat.get("line", ""), bool(cat.get("histology", False))

    def _force_myringotomy_plan_with_adenotomy(self, ops_df: pd.DataFrame) -> pd.DataFrame:
        """Если в тот же день у того же КВС есть аденотомия — миринготомия → план."""
        if ops_df is None or ops_df.empty:
            return ops_df
        df = ops_df.copy()
        df["_day"] = pd.to_datetime(df["Дата"], errors="coerce").dt.normalize()
        adenotomy_keys = set(
            zip(
                df.loc[df["Категория"] == "Аденотомия", "КВС"].astype(str),
                df.loc[df["Категория"] == "Аденотомия", "_day"],
            )
        )
        if not adenotomy_keys:
            return df.drop(columns=["_day"])

        mask = df["Категория"].isin(["Миринготомия экстр", "Миринготомия план"])
        for i in df.loc[mask].index:
            key = (str(df.at[i, "КВС"]), df.at[i, "_day"])
            if key in adenotomy_keys:
                df.at[i, "Категория"] = "Миринготомия план"
        return df.drop(columns=["_day"])

    def _extract_surgeon(self, team_str: str) -> str:
        match = re.search(r"(?:Основной\s+)?Хирург\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]?)", team_str)
        if match:
            return match.group(1)
        return "Не указан"

    def summary_table(self, ops_df: pd.DataFrame, summary_cfg: dict):
        return build_summary_tables(ops_df, summary_cfg, self.categories)


def build_summary_tables(ops_df: pd.DataFrame, summary_cfg: dict, categories: List[dict] = None):
    """Агрегация категория × неделя; план/экстренно — по спискам шаблона."""
    if ops_df is None or ops_df.empty:
        return pd.DataFrame(), pd.DataFrame(), []

    ops_df = ops_df.copy()
    ops_df["Неделя_начало"] = ops_df["Дата"].apply(
        lambda d: d - pd.Timedelta(days=d.weekday()) if pd.notna(d) else pd.NaT
    )
    weeks_sorted = sorted(ops_df["Неделя_начало"].dropna().unique())

    cat_order = list(summary_cfg.get("category_rows", {}).keys())
    if not cat_order and categories:
        cat_order = [c["category"] for c in categories]

    cat_table = pd.DataFrame(0, index=cat_order, columns=weeks_sorted)
    for _, r in ops_df.iterrows():
        cat = r["Категория"]
        week = r["Неделя_начало"]
        if cat in cat_table.index and pd.notna(week):
            cat_table.loc[cat, week] += 1

    emerg_set = set(summary_cfg.get("emergency_categories", []))
    plan_set = set(summary_cfg.get("plan_categories", []))
    emerg_idx = [c for c in cat_table.index if c in emerg_set]
    plan_idx = [c for c in cat_table.index if c in plan_set]

    total_ops = cat_table.sum(axis=0)
    emerg_counts = cat_table.loc[emerg_idx].sum(axis=0) if emerg_idx else pd.Series(0, index=weeks_sorted)
    plan_counts = cat_table.loc[plan_idx].sum(axis=0) if plan_idx else pd.Series(0, index=weeks_sorted)

    children_ops = ops_df[ops_df["Возраст"].fillna(99) < 18]
    children_counts = children_ops.groupby("Неделя_начало")["КВС"].nunique()
    unique_patients = ops_df.groupby("Неделя_начало")["КВС"].nunique()

    totals_df = pd.DataFrame(
        index=[
            "Всего операций",
            "Экстренно операций",
            "План операций",
            "Дети всего",
            "Взрослые",
            "Человек",
        ],
        columns=weeks_sorted,
    )
    totals_df.loc["Всего операций"] = total_ops.reindex(weeks_sorted, fill_value=0).values
    totals_df.loc["Экстренно операций"] = emerg_counts.reindex(weeks_sorted, fill_value=0).values
    totals_df.loc["План операций"] = plan_counts.reindex(weeks_sorted, fill_value=0).values
    totals_df.loc["Дети всего"] = children_counts.reindex(weeks_sorted, fill_value=0).values
    totals_df.loc["Человек"] = unique_patients.reindex(weeks_sorted, fill_value=0).values
    totals_df.loc["Взрослые"] = (
        totals_df.loc["Человек"].astype(int) - totals_df.loc["Дети всего"].astype(int)
    ).clip(lower=0).values

    return cat_table.astype(int), totals_df.astype(int), weeks_sorted
