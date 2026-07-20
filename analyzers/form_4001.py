# analyzers/form_4001.py
"""
Форма «ХИРУРГИЧЕСКАЯ РАБОТА ОРГАНИЗАЦИИ» (4001) — как в шаблоне сводной.

Структура листа (Апрель и др.), колонки L–S:
  L  Наименование операции
  M  № строки
  N  гр.3  всего
  O  гр.4  0–14 лет
  P  гр.5  до 1 года
  Q  гр.6  15–17 лет
  R  гр.28 морфология
  S  гр.3  «лица старше трудоспособного» / Всего (из гр.3 т.4000)

Итоги N/O/R по строкам 5.1 / 5.2 / 6 / 6.1 / 17 в шаблоне — формулы от H
(ИТОГ категорий слева). Их не затираем: пишем только «ручные» ячейки
(P, Q, S, пенсионная строка) и оставляем формулы SUM.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

import pandas as pd

# Точные подписи из шаблона (в т.ч. пробел в начале у 5.2)
FORM_ROWS: List[dict] = [
    {"line": "5", "name": "операции на органах уха, горла, носа", "excel_row": 11, "kind": "parent"},
    {"line": "5.1", "name": "из них - на ухе", "excel_row": 12, "kind": "leaf"},
    {"line": "5.2", "name": " на миндалинах и аденоидах", "excel_row": 13, "kind": "leaf"},
    {"line": "6", "name": "операции на органах дыхания", "excel_row": 14, "kind": "leaf"},
    {"line": "6.1", "name": "из них - на трахее", "excel_row": 15, "kind": "leaf"},
    {"line": "17", "name": "операции на коже и подкожной клетчатке", "excel_row": 16, "kind": "leaf"},
    {"line": "", "name": "Операции лицам пенсионного возраста", "excel_row": 17, "kind": "pension"},
    {"line": "", "name": "Всего операций", "excel_row": 18, "kind": "total"},
    {"line": "", "name": "Эндоскопических всего", "excel_row": 20, "kind": "endo"},
]

# Соответствие формул шаблона: строка формы → категории (колонка H / B)
# N12 = H6+H9+H16+H22+H26+H24+H32
LINE_TOTAL_CATS: Dict[str, List[str]] = {
    "5.1": [
        "Миринготомия план",
        "Миринготомия экстр",
        "Фурункул НСП",
        "Рубцы мочки уха",
        "Удаление новообр уха",
        "Антромастоидотомия",
        "Субперостальный абсцесс за ухом",
    ],
    # N13 = H4+H5+H8+H13
    "5.2": ["Аденотомия", "Тонзиллотомия", "ПТА", "Тозиллэктомия"],
    # N14 = длинная сумма дыхания
    "6": [
        "Пластика раковин",
        "Полипотомия",
        "Септопластика",
        "Гайморотомия",
        "Увулопластика",
        "Пластика нёба",
        "Фурункул носа",
        "Трахеостомия",
        "Репозиция костей носа",
        "Флегмона шеи",
        "Остановка кров",
        "Синехии нос",
        "Удаление новообр носа",
        "Удаление новообр глотки",
        "Удаление новообр гортани",
        "Удаление инородного тела",
        "Фронтотомия",
        "Ревизия п/о полости",
        "Заглоточный абсцесс",
        "Биопсия гортани ",
    ],
    # N15 = H18
    "6.1": ["Трахеостомия"],
    # N16 = H35+H33+H29
    "17": ["Наложение вторичных швов", "Пластика местными тканями", "ПХО"],
}

# O12 = H6 − Q12  (только миринготомия план)
O_BASE_CATS: Dict[str, List[str]] = {
    "5.1": ["Миринготомия план"],
    # O13 = H4+H5 − Q13
    "5.2": ["Аденотомия", "Тонзиллотомия"],
}

# R12 = H22+H26; R13 = H4+H5+H13; R14 = H10+H12+H14+H27+H28+H37
LINE_HIST_CATS: Dict[str, List[str]] = {
    "5.1": ["Рубцы мочки уха", "Удаление новообр уха"],
    "5.2": ["Аденотомия", "Тонзиллотомия", "Тозиллэктомия"],
    "6": [
        "Полипотомия",
        "Гайморотомия",
        "Увулопластика",
        "Удаление новообр глотки",
        "Удаление новообр гортани",
        "Биопсия гортани ",
    ],
    "6.1": [],
    "17": [],
}

# R11 = R12+R13+H25 (удаление новообр носа добавляется к ЛОР-блоку)
R11_EXTRA = ["Удаление новообр носа"]

# N20 эндоскопия = H4+H7+H10+H11+H12+H13+H28+H37
ENDO_CATS = [
    "Аденотомия",
    "Пластика раковин",
    "Полипотомия",
    "Септопластика",
    "Гайморотомия",
    "Тозиллэктомия",
    "Удаление новообр гортани",
    "Биопсия гортани ",
]

FORM_LINE_LABELS = {r["line"]: r["name"] for r in FORM_ROWS if r["line"]}


def resolve_line_total_cats(form_cfg: Optional[dict] = None) -> Dict[str, List[str]]:
    """Списки категорий для N: из form_4001.line_categories или встроенные LINE_TOTAL_CATS."""
    base = {k: list(v) for k, v in LINE_TOTAL_CATS.items()}
    extra = (form_cfg or {}).get("line_categories") or {}
    for line, cats in extra.items():
        key = str(line)
        merged = list(base.get(key, []))
        for c in cats or []:
            s = str(c)
            if s not in merged:
                merged.append(s)
        base[key] = merged
    return base


def resolve_line_hist_cats(form_cfg: Optional[dict] = None) -> Dict[str, List[str]]:
    base = {k: list(v) for k, v in LINE_HIST_CATS.items()}
    extra = (form_cfg or {}).get("hist_categories") or {}
    for line, cats in extra.items():
        key = str(line)
        merged = list(base.get(key, []))
        for c in cats or []:
            s = str(c)
            if s not in merged:
                merged.append(s)
        base[key] = merged
    return base


def resolve_endo_cats(form_cfg: Optional[dict] = None) -> List[str]:
    extra = (form_cfg or {}).get("endo_categories")
    if extra is None:
        return list(ENDO_CATS)
    merged = list(ENDO_CATS)
    for c in extra:
        s = str(c)
        if s not in merged:
            merged.append(s)
    return merged


def _cat_counts(month_ops: pd.DataFrame) -> Dict[str, int]:
    if month_ops is None or month_ops.empty or "Категория" not in month_ops.columns:
        return {}
    return {str(k): int(v) for k, v in month_ops["Категория"].value_counts().items()}


def _sum_cats(counts: Dict[str, int], cats: Sequence[str]) -> int:
    return int(sum(counts.get(c, 0) for c in cats))


def _age_mask(ops: pd.DataFrame, lo: Optional[float], hi: Optional[float]) -> pd.Series:
    age = pd.to_numeric(ops.get("Возраст"), errors="coerce")
    m = age.notna()
    if lo is not None:
        m &= age >= lo
    if hi is not None:
        m &= age <= hi
    return m


def _count_ops_in_cats(ops: pd.DataFrame, cats: Sequence[str], mask: Optional[pd.Series] = None) -> int:
    if ops is None or ops.empty or not cats:
        return 0
    m = ops["Категория"].isin(list(cats))
    if mask is not None:
        m &= mask
    return int(m.sum())


def compute_form_4001(
    month_ops: pd.DataFrame,
    categories: List[dict] = None,  # noqa: ARG001 — совместимость вызовов
    line_rows: Dict[str, int] = None,  # noqa: ARG001
    pension_age: int = 60,
    form_cfg: Optional[dict] = None,
) -> Dict[str, Any]:
    """
    Счётчики «как после пересчёта Excel»: N/R по формулам от категорий,
    Q/P/S — по возрасту внутри групп строки.
    """
    ops = month_ops if month_ops is not None else pd.DataFrame()
    counts = _cat_counts(ops)
    ages = pd.to_numeric(ops["Возраст"], errors="coerce") if not ops.empty else pd.Series(dtype=float)

    line_totals = resolve_line_total_cats(form_cfg)
    line_hist = resolve_line_hist_cats(form_cfg)
    endo_cats = resolve_endo_cats(form_cfg)

    lines: Dict[str, dict] = {}
    for line, cats in line_totals.items():
        total = _sum_cats(counts, cats)
        hist = _sum_cats(counts, line_hist.get(line, []))
        # 15–17 и до 1 года — по всем операциям группы строки
        m_u1 = _age_mask(ops, None, 0.999) if not ops.empty else None
        m_1517 = _age_mask(ops, 15, 17) if not ops.empty else None
        m_pension = (ages >= pension_age) if not ops.empty else None
        age_under_1 = _count_ops_in_cats(ops, cats, m_u1)
        age_15_17 = _count_ops_in_cats(ops, cats, m_1517)
        senior = _count_ops_in_cats(ops, cats, m_pension)

        # O как в шаблоне:
        # 5.1 / 5.2 — формула «база категорий − Q»
        # 6 / 6.1 / 17 — в шаблоне константа 0 (не возрастная выборка)
        if line in O_BASE_CATS:
            base = _sum_cats(counts, O_BASE_CATS[line])
            age_0_14 = max(0, base - age_15_17)
        else:
            age_0_14 = 0

        lines[line] = {
            "total": total,
            "age_0_14": age_0_14,
            "age_under_1": age_under_1,
            "age_15_17": age_15_17,
            "histology": hist,
            "senior": senior,
        }

    # родитель 5
    b51, b52 = lines["5.1"], lines["5.2"]
    r11_extra = _sum_cats(counts, R11_EXTRA)
    lines["5"] = {
        "total": b51["total"] + b52["total"],
        "age_0_14": b51["age_0_14"] + b52["age_0_14"],
        "age_under_1": b51["age_under_1"] + b52["age_under_1"],
        "age_15_17": b51["age_15_17"] + b52["age_15_17"],
        "histology": b51["histology"] + b52["histology"] + r11_extra,
        "senior": b51["senior"] + b52["senior"],
    }

    pension = int((ages >= pension_age).sum()) if not ops.empty else 0
    endo = _sum_cats(counts, endo_cats)

    # Всего: N18 = N16+N14+N11; O18 = O11+O14+O16−Q14; …
    b6, b17 = lines["6"], lines["17"]
    parent = lines["5"]
    total_row = {
        "total": parent["total"] + b6["total"] + b17["total"],
        "age_0_14": parent["age_0_14"] + b6["age_0_14"] + b17["age_0_14"] - b6["age_15_17"],
        "age_under_1": parent["age_under_1"] + b6["age_under_1"] + b17["age_under_1"],
        "age_15_17": parent["age_15_17"] + b6["age_15_17"] + b17["age_15_17"],
        "histology": parent["histology"] + b6["histology"] + b17["histology"],
        "senior": parent["senior"] + b6["senior"] + b17["senior"],
    }
    # O18 в шаблоне: SUM(O11+O14+O16)-Q14 — без вычитания Q из O11/O16
    total_row["age_0_14"] = parent["age_0_14"] + b6["age_0_14"] + b17["age_0_14"] - b6["age_15_17"]

    return {
        "lines": lines,
        "pension": pension,
        "endo": endo,
        "total": total_row,
    }


def form_4001_preview_rows(stats: dict) -> List[dict]:
    """Строки превью — в точности порядок и подписи шаблона."""
    lines = stats.get("lines") or {}
    empty = {
        "total": 0,
        "age_0_14": 0,
        "age_under_1": 0,
        "age_15_17": 0,
        "histology": 0,
        "senior": 0,
    }
    rows: List[dict] = []
    for spec in FORM_ROWS:
        kind = spec["kind"]
        line = spec["line"]
        if kind == "pension":
            rows.append(
                {
                    "name": spec["name"],
                    "line": "",
                    "total": int(stats.get("pension", 0)),
                    "age_0_14": 0,
                    "age_under_1": 0,
                    "age_15_17": 0,
                    "histology": 0,
                    "senior": "",
                }
            )
        elif kind == "total":
            b = stats.get("total") or empty
            rows.append({"name": spec["name"], "line": "", **{k: b.get(k, 0) for k in empty}})
        elif kind == "endo":
            rows.append(
                {
                    "name": spec["name"],
                    "line": "",
                    "total": int(stats.get("endo", 0)),
                    "age_0_14": "",
                    "age_under_1": "",
                    "age_15_17": "",
                    "histology": "",
                    "senior": "",
                }
            )
        else:
            b = lines.get(line) or empty
            rows.append(
                {
                    "name": spec["name"],
                    "line": line,
                    "total": b.get("total", 0),
                    "age_0_14": b.get("age_0_14", 0),
                    "age_under_1": b.get("age_under_1", 0),
                    "age_15_17": b.get("age_15_17", 0),
                    "histology": b.get("histology", 0),
                    "senior": b.get("senior", 0),
                }
            )
    return rows


def write_form_4001(
    ws,
    month_ops: pd.DataFrame,
    categories: List[dict],
    form_cfg: dict,
    pension_age: int = 60,
) -> dict:
    """
    Не затирает формулы SUM(H…) / SUM(N12:N13) и т.п.
    Пишет только ячейки, которые в шаблоне — числа (P, Q, S, пенсия).
    N/O/R листовых строк с формулами остаются — Excel пересчитает от H.
    """
    if not form_cfg or not form_cfg.get("enabled", True):
        return {"written": 0}

    cols = form_cfg.get("cols") or {}
    col_total = int(cols.get("total", 14))       # N
    col_014 = int(cols.get("age_0_14", 15))      # O
    col_u1 = int(cols.get("age_under_1", 16))   # P
    col_1517 = int(cols.get("age_15_17", 17))   # Q
    col_hist = int(cols.get("histology", 18))   # R
    col_senior = int(cols.get("senior", 19))     # S
    pension_row = int(form_cfg.get("pension_row", 17))

    stats = compute_form_4001(
        month_ops, categories, pension_age=pension_age, form_cfg=form_cfg
    )
    written = 0
    lines = stats["lines"]

    def _is_formula(cell) -> bool:
        v = cell.value
        return isinstance(v, str) and v.startswith("=")

    # Листовые строки 5.1 / 5.2 / 6 / 6.1 / 17
    for spec in FORM_ROWS:
        if spec["kind"] != "leaf":
            continue
        line = spec["line"]
        row = int(spec["excel_row"])
        b = lines.get(line) or {}

        # P — до 1 года (в шаблоне пусто/0)
        cell_p = ws.cell(row, col_u1)
        if not _is_formula(cell_p):
            cell_p.value = int(b.get("age_under_1", 0))
            written += 1

        # Q — 15–17 (в шаблоне часто ручное число; нужно для формул O12/O13)
        cell_q = ws.cell(row, col_1517)
        if not _is_formula(cell_q):
            cell_q.value = int(b.get("age_15_17", 0))
            written += 1

        # S — старше трудоспособного (в шаблоне часто ручные числа т.4000)
        cell_s = ws.cell(row, col_senior)
        if not _is_formula(cell_s):
            cell_s.value = int(b.get("senior", 0))
            written += 1

        # O: только если в шаблоне НЕ формула и это не «константа 0» у дыхания/кожи —
        # для 6/6.1/17 оставляем 0 как в шаблоне; для 5.1/5.2 O — формула, не трогаем
        cell_o = ws.cell(row, col_014)
        if not _is_formula(cell_o) and line not in ("6", "6.1", "17"):
            cell_o.value = int(b.get("age_0_14", 0))
            written += 1

        # R без формулы (6.1, 17) — гистология
        cell_r = ws.cell(row, col_hist)
        if not _is_formula(cell_r):
            cell_r.value = int(b.get("histology", 0))
            written += 1

        # N с формулой SUM(H…) не трогаем
        cell_n = ws.cell(row, col_total)
        if not _is_formula(cell_n):
            cell_n.value = int(b.get("total", 0))
            written += 1

    # Родитель 5: S11 часто число; N/O/Q/R — формулы
    parent = next(r for r in FORM_ROWS if r["line"] == "5")
    prow = int(parent["excel_row"])
    cell_s = ws.cell(prow, col_senior)
    if not _is_formula(cell_s):
        cell_s.value = int((lines.get("5") or {}).get("senior", 0))
        written += 1
    # R11 в шаблоне = R12+R13+H25 — оставляем формулу; если была испорчена — восстановим
    cell_r11 = ws.cell(prow, col_hist)
    if _is_formula(cell_r11) and "H25" in str(cell_r11.value):
        pass  # штатная формула шаблона
    elif not _is_formula(cell_r11) or "H25" not in str(cell_r11.value or ""):
        # восстановить канон шаблона
        r51 = next(r for r in FORM_ROWS if r["line"] == "5.1")["excel_row"]
        r52 = next(r for r in FORM_ROWS if r["line"] == "5.2")["excel_row"]
        cell_r11.value = f"=SUM(R{r51}+R{r52}+H25)"
        written += 1

    # Пенсионная строка
    ws.cell(pension_row, col_total).value = int(stats.get("pension", 0))
    ws.cell(pension_row, col_014).value = 0
    cell_pr = ws.cell(pension_row, col_hist)
    if not _is_formula(cell_pr):
        cell_pr.value = 0
    written += 3

    return {"written": written, "stats": stats}
