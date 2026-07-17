# analyzers/problem_codes.py
"""Сводка неклассифицированных кодов + черновик для config.yaml."""
from __future__ import annotations

from typing import List

import pandas as pd

from analyzers.ksg_catalog import get_catalog


def build_problem_codes_table(ops: pd.DataFrame) -> pd.DataFrame:
    """
    Уникальные коды среди «Не классифицировано» с числом операций,
    подсказкой КСГ и фрагментом для config.yaml.
    """
    if ops is None or ops.empty:
        return pd.DataFrame(
            columns=["Код", "Операций", "Услуга_пример", "Подсказка_КСГ", "Черновик_config"]
        )
    uncl = ops[ops["Категория"] == "Не классифицировано"].copy()
    if uncl.empty:
        return pd.DataFrame(
            columns=["Код", "Операций", "Услуга_пример", "Подсказка_КСГ", "Черновик_config"]
        )

    cat = get_catalog()
    rows: List[dict] = []
    codes = uncl["Код"].fillna("").astype(str).str.strip()
    for code, grp in uncl.groupby(codes):
        if not code or code.lower() in ("nan", "none"):
            code_label = "(без кода)"
            draft = (
                "  # операция без кода — добавьте name_keywords по названию услуги\n"
                '  # - category: "…"\n'
                "  #   codes: []\n"
                '  #   name_keywords: ["…"]'
            )
            hint = ""
        else:
            code_label = code
            hint = cat.hint_for(code)
            draft = cat.suggest_config_snippet(code)
        sample = ""
        if "Услуга" in grp.columns and len(grp):
            sample = str(grp.iloc[0].get("Услуга", "") or "")[:120]
        rows.append(
            {
                "Код": code_label,
                "Операций": int(len(grp)),
                "Услуга_пример": sample,
                "Подсказка_КСГ": hint,
                "Черновик_config": draft,
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Операций", "Код"], ascending=[False, True]).reset_index(drop=True)
    return df


def format_config_draft(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "# нет неклассифицированных кодов\n"
    parts = [
        "# Черновик для surgery_categories в config.yaml",
        "# Проверьте category / group / line и перенесите в config.yaml",
        "",
    ]
    for _, row in df.iterrows():
        parts.append(f"# --- {row['Код']} ({row['Операций']} опер.) ---")
        if row.get("Подсказка_КСГ"):
            parts.append(f"# КСГ: {row['Подсказка_КСГ']}")
        if row.get("Услуга_пример"):
            parts.append(f"# пример: {row['Услуга_пример']}")
        parts.append(str(row.get("Черновик_config") or ""))
        parts.append("")
    return "\n".join(parts)
