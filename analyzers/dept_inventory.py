# analyzers/dept_inventory.py
"""Инвентаризация операций отделения из op-журналов и сопоставление с KSG."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from analyzers.category_registry import suggest_keywords_from_name
from analyzers.dept_config import DEPT_REPORT_SOURCES
from analyzers.io_utils import smart_read_excel
from analyzers.ksg_catalog import get_catalog

_CODE_RE = re.compile(r"A\d{2}\.\d{2}\.\d{3}(?:\.\d{3})?")


def extract_codes_from_report(report_path: str | Path, department: str) -> pd.DataFrame:
    """
    Уникальные коды услуг в отчёте отделения с частотой и примером названия.
    Колонки: Код, Операций, Услуга_пример, Дата_мин, Дата_макс.
    """
    df = smart_read_excel(str(report_path))
    dept_col = "Отделение госпитализации"
    mask = df[dept_col].astype(str).str.contains(department, na=False) | df["Оперблок"].astype(
        str
    ).str.contains(department, na=False)
    sub = df[mask].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Код", "Операций", "Услуга_пример", "Дата_мин", "Дата_макс"])

    date_col = None
    for col in sub.columns:
        if "дата" in str(col).lower() and "начала" in str(col).lower():
            date_col = col
            break
    if date_col:
        sub["_dt"] = pd.to_datetime(sub[date_col], dayfirst=True, errors="coerce")
    else:
        sub["_dt"] = pd.NaT

    rows: List[dict] = []
    for _, row in sub.iterrows():
        text = str(row.get("Услуга", "") or "")
        codes = _CODE_RE.findall(text)
        if not codes:
            continue
        for code in codes:
            rows.append({"Код": code, "Услуга": text, "_dt": row["_dt"]})

    if not rows:
        return pd.DataFrame(columns=["Код", "Операций", "Услуга_пример", "Дата_мин", "Дата_макс"])

    tmp = pd.DataFrame(rows)
    out: List[dict] = []
    for code, grp in tmp.groupby("Код"):
        sample = str(grp.iloc[0]["Услуга"] or "")[:160]
        dts = grp["_dt"].dropna()
        out.append(
            {
                "Код": code,
                "Операций": int(len(grp)),
                "Услуга_пример": sample,
                "Дата_мин": dts.min().date().isoformat() if not dts.empty else "",
                "Дата_макс": dts.max().date().isoformat() if not dts.empty else "",
            }
        )
    result = pd.DataFrame(out).sort_values(["Операций", "Код"], ascending=[False, True])
    return result.reset_index(drop=True)


def enrich_with_ksg(codes_df: pd.DataFrame) -> pd.DataFrame:
    """Добавляет KSG-поля из KSGoperacii.csv."""
    if codes_df is None or codes_df.empty:
        return codes_df
    cat = get_catalog()
    ksgs: List[str] = []
    names: List[str] = []
    hints: List[str] = []
    in_catalog: List[bool] = []
    for _, row in codes_df.iterrows():
        code = str(row.get("Код") or "")
        info = cat.lookup(code)
        if info:
            names.append(info.get("name") or "")
            ksgs.append(", ".join(info.get("ksg") or []))
            hints.append(cat.hint_for(code))
            in_catalog.append(True)
        else:
            names.append("")
            ksgs.append("")
            hints.append("нет в KSGoperacii.csv")
            in_catalog.append(False)
    out = codes_df.copy()
    out["Наименование_КСГ"] = names
    out["КСГ"] = ksgs
    out["Подсказка"] = hints
    out["В_справочнике"] = in_catalog
    return out


def build_inventory_table(report_path: str | Path, department: str) -> pd.DataFrame:
    base = extract_codes_from_report(report_path, department)
    return enrich_with_ksg(base)


_CODE_PREFIX_RE = re.compile(r"^A\d{2}\.\d{2}\.\d{3}(?:\.\d{3})?\s*[-–—:]?\s*", re.I)


def strip_service_code_prefix(text: str) -> str:
    """Убирает префикс кода A16… из названия услуги."""
    s = str(text or "").strip()
    return _CODE_PREFIX_RE.sub("", s).strip() or s


def shorten_category_label(name: str, max_len: int = 62) -> str:
    """
    Короткие подписи для Excel: убираем «Широкое», обрезаем «компонентом»
    и ограничиваем длину по словам.
    """
    s = str(name or "").strip()
    if not s:
        return s
    # частный случай из отчётов хирургии
    s = s.replace(
        "Широкое иссечение новообразования кожи с реконструктивно-пластическим компонентом",
        "Иссечение новообразования кожи с реконструктивно-пластическим",
    )
    s = s.replace("Широкое иссечение ", "Иссечение ")
    s = s.replace(" с реконструктивно-пластическим компонентом", " с реконстр.-пласт. компонентом")
    if len(s) <= max_len:
        return s
    # обрезка по границе слова
    cut = s[: max_len - 1]
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0]
    return cut.rstrip(".,;:-") + "…"


def category_display_name(
    *,
    ksg_name: str = "",
    sample: str = "",
    code: str = "",
    used: Optional[set] = None,
) -> str:
    """
    Единый формат имени категории — без кода A16… в тексте.
    При коллизии имён добавляет короткий суффикс (…2, …3).
    """
    used = used if used is not None else set()
    base = (ksg_name or "").strip()
    if not base:
        base = strip_service_code_prefix(sample.split("(")[0] if sample else "")
    if not base:
        base = code or "Операция"
    base = shorten_category_label(base)
    name = base
    if name not in used:
        used.add(name)
        return name
    i = 2
    while True:
        cand = f"{base} ({i})"
        if len(cand) > 60:
            cand = f"{base[:50]}… ({i})"
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def build_categories_from_codes(
    codes_df: pd.DataFrame,
    *,
    group: str = "операции",
    line: str = "6",
) -> Tuple[List[dict], Dict[str, int], dict]:
    """
    1 код = 1 категория, все plan (до классификации по ЭМК).
    Имена без кода услуги в тексте.
    Возвращает (surgery_categories, category_rows, summary_meta).
    """
    if codes_df is None or codes_df.empty:
        return [], {}, {}

    sorted_df = codes_df.sort_values(["Операций", "Код"], ascending=[False, True])
    used_names: set = set()
    categories: List[dict] = []
    category_rows: Dict[str, int] = {}
    plan_names: List[str] = []

    for i, (_, row) in enumerate(sorted_df.iterrows()):
        code = str(row["Код"]).strip()
        ksg_name = str(row.get("Наименование_КСГ") or "").strip()
        sample = str(row.get("Услуга_пример") or "").strip()
        cat_name = category_display_name(
            ksg_name=ksg_name, sample=sample, code=code, used=used_names
        )
        keywords = suggest_keywords_from_name(ksg_name or strip_service_code_prefix(sample) or cat_name)
        if not keywords and code:
            keywords = [code.lower()]
        categories.append(
            {
                "category": cat_name,
                "codes": [code],
                "group": group,
                "line": line,
                "histology": False,
                "name_keywords": keywords,
            }
        )
        excel_row = 4 + i
        category_rows[cat_name] = excel_row
        plan_names.append(cat_name)

    n = len(categories)
    last_cat = 4 + n - 1 if n else 3
    blank = last_cat + 1
    total_row = blank + 1
    summary_meta = {
        "category_rows": category_rows,
        "plan_categories": plan_names,
        "emergency_categories": [],
        "totals_rows": {
            "total": total_row,
            "emergency": total_row + 1,
            "plan": total_row + 2,
            "children": total_row + 4,
            "adults": total_row + 5,
            "patients": total_row + 6,
        },
        "children": total_row + 4,
        "patients": total_row + 6,
    }
    return categories, category_rows, summary_meta


def build_summary_cfg_draft(
    summary_key: str,
    categories: List[dict],
    summary_meta: dict,
    *,
    year: int = 2026,
    backup_keep: int = 20,
) -> dict:
    from analyzers.dept_config import DEPT_REPORT_SOURCES, default_summary_filename

    meta = DEPT_REPORT_SOURCES.get(summary_key) or {}
    default_path = meta.get("default_path") or "Операции сводная {year}.xlsx"
    default_path = str(default_path).format(year=year)
    totals = summary_meta.get("totals_rows") or {}
    return {
        "default_path": default_path,
        "backup_keep": backup_keep,
        "year": year,
        "sheet_names": {
            1: "Январь",
            2: "Февраль",
            3: "Март",
            4: "Апрель",
            5: "Май",
            6: "Июнь",
            7: "Июль",
            8: "Август",
            9: "Сентябрь",
            10: "Октябрь ",
            11: "Ноябрь",
            12: "Декабрь",
        },
        "category_rows": dict(summary_meta.get("category_rows") or {}),
        "plan_categories": list(summary_meta.get("plan_categories") or []),
        "emergency_categories": [],
        "totals_rows": {
            "children": int(totals.get("children") or totals.get("patients", 0) - 2 or 0),
            "patients": int(totals.get("patients") or 0),
        },
    }


def inventory_from_source(
    summary_key: str,
    reports_dir: str | Path,
) -> Tuple[pd.DataFrame, List[dict], dict]:
    """Полный цикл: отчёт → таблица → categories + summary_meta."""
    meta = DEPT_REPORT_SOURCES.get(summary_key)
    if not meta:
        raise ValueError(f"Неизвестный summary_key: {summary_key}")
    path = Path(reports_dir) / meta["report_file"]
    table = build_inventory_table(path, meta["department"])
    cats, _rows, smeta = build_categories_from_codes(table)
    return table, cats, smeta


def export_inventory_excel(
    table: pd.DataFrame,
    output_path: str | Path,
    *,
    summary_key: str = "",
) -> Path:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        table.to_excel(writer, sheet_name="Все коды", index=False)
        if not table.empty:
            missing = table[~table["В_справочнике"].astype(bool)]
            missing.to_excel(writer, sheet_name="Без KSG", index=False)
            table.head(20).to_excel(writer, sheet_name="Топ-20", index=False)
        if summary_key:
            pd.DataFrame({"summary_key": [summary_key]}).to_excel(
                writer, sheet_name="Meta", index=False
            )
    return out


def format_yaml_categories_snippet(categories: List[dict], limit: int = 5) -> str:
    """Краткий YAML-фрагмент для ручной проверки."""
    lines = ["# surgery_categories (фрагмент):"]
    for cat in categories[:limit]:
        lines.append(f"  - category: \"{cat['category']}\"")
        lines.append(f"    codes: {cat['codes']}")
    if len(categories) > limit:
        lines.append(f"  # … ещё {len(categories) - limit} категорий")
    return "\n".join(lines)
