# analyzers/form14_overrides.py
"""Ручные переназначения A16 → строка ФСН 14 (калибровка хирургами)."""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import yaml

DEFAULT_FILENAME = "form14_overrides.yaml"
MANUAL_COL = "Строка_ФСН14_ручная"
COMMENT_COL = "Комментарий"
HISTOLOGY_COL = "Морфология"

# Excel трактует ячейки, начинающиеся с этих символов, как формулы (CVE-style injection).
_EXCEL_FORMULA_LEADERS = frozenset("=+-@\t\r")


def neutralize_excel_formula(value: Any) -> Any:
    """
    Превращает потенциальную формулу Excel в безопасный текст.
    Префикс "'" — стандартный способ заставить Excel показать строку как есть.
    """
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, float) and pd.isna(value):
        return value
    s = str(value)
    if not s:
        return s
    # уже нейтрализовано ранее
    if s.startswith("'"):
        return s
    lead = s.lstrip()[:1]
    if lead and lead in _EXCEL_FORMULA_LEADERS:
        return "'" + s
    return s


def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Нейтрализует формулы во всех текстовых ячейках DataFrame перед to_excel."""
    if df is None or getattr(df, "empty", True):
        return df
    out = df.copy()
    for col in out.columns:
        series = out[col]
        if series.dtype == object or pd.api.types.is_string_dtype(series):
            out[col] = series.map(
                lambda v: neutralize_excel_formula(v) if isinstance(v, str) else v
            )
    return out


def default_path(app_dir: Path | str) -> Path:
    return Path(app_dir) / DEFAULT_FILENAME


def empty_store() -> dict:
    return {"by_code": {}, "by_category": {}}


def load_overrides(path: Path | str) -> dict:
    p = Path(path)
    if not p.exists():
        return empty_store()
    data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    store = empty_store()
    by_code = data.get("by_code") or {}
    by_cat = data.get("by_category") or {}
    if isinstance(by_code, dict):
        store["by_code"] = {str(k).strip(): _norm_entry(v) for k, v in by_code.items() if str(k).strip()}
    if isinstance(by_cat, dict):
        store["by_category"] = {
            str(k).strip(): _norm_entry(v) for k, v in by_cat.items() if str(k).strip()
        }
    return store


def _norm_line(raw: Any) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, float) and raw == int(raw):
        return str(int(raw))
    if isinstance(raw, int):
        return str(raw)
    s = str(raw).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    # "17.0" from Excel
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except ValueError:
        pass
    return s


def _norm_histology(raw: Any) -> Optional[bool]:
    """None = не задано; True/False = явный флаг морфологии."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, bool):
        return raw
    s = str(raw).strip().lower()
    if s in ("", "nan", "none", "-", "null"):
        return None
    if s in ("1", "true", "yes", "да", "y", "д"):
        return True
    if s in ("0", "false", "no", "нет", "n", "н"):
        return False
    return None


def _safe_comment(raw: Any) -> str:
    return str(neutralize_excel_formula(str(raw or "").strip()) or "")


def _norm_entry(v: Any) -> dict:
    if isinstance(v, dict):
        return {
            "line": _norm_line(v.get("line")),
            "comment": _safe_comment(v.get("comment")),
            "by": str(v.get("by") or "").strip(),
            "at": str(v.get("at") or "").strip(),
            "histology": _norm_histology(v.get("histology")),
        }
    return {"line": _norm_line(v), "comment": "", "by": "", "at": "", "histology": None}


def _entry_worth_keeping(e: dict) -> bool:
    return bool(e.get("line")) or e.get("histology") is not None


def _dump_entry(e: dict) -> dict:
    out = {
        "line": e.get("line") or "",
        "comment": e.get("comment") or "",
        "by": e.get("by") or "",
        "at": e.get("at") or "",
    }
    if e.get("histology") is not None:
        out["histology"] = bool(e["histology"])
    return out


def save_overrides(store: dict, path: Path | str) -> Path:
    p = Path(path)
    out = empty_store()
    for k, v in (store.get("by_code") or {}).items():
        e = _norm_entry(v)
        if _entry_worth_keeping(e):
            out["by_code"][str(k).strip()] = _dump_entry(e)
    for k, v in (store.get("by_category") or {}).items():
        e = _norm_entry(v)
        if _entry_worth_keeping(e):
            out["by_category"][str(k).strip()] = _dump_entry(e)
    text = yaml.safe_dump(out, allow_unicode=True, default_flow_style=False, sort_keys=False)
    p.write_text(text, encoding="utf-8")
    return p


def lookup_entry(
    store: dict,
    *,
    code: str = "",
    category: str = "",
) -> Optional[dict]:
    """Любая запись (со строкой и/или морфологией). Сначала код, затем категория."""
    code = str(code or "").strip()
    category = str(category or "").strip()
    by_code = store.get("by_code") or {}
    by_cat = store.get("by_category") or {}
    if code and code in by_code:
        return _norm_entry(by_code[code])
    if category:
        if category in by_cat:
            return _norm_entry(by_cat[category])
        stripped = category.strip()
        if stripped in by_cat:
            return _norm_entry(by_cat[stripped])
    return None


def lookup_override(
    store: dict,
    *,
    code: str = "",
    category: str = "",
) -> Optional[Tuple[str, dict]]:
    """Возвращает (line, entry) или None, если ручная строка формы не задана."""
    e = lookup_entry(store, code=code, category=category)
    if e and e.get("line"):
        return e["line"], e
    return None


def histology_for(store: dict, *, code: str = "", category: str = "") -> Optional[bool]:
    """Явный флаг морфологии из overrides или None (брать из категории/операции)."""
    e = lookup_entry(store, code=code, category=category)
    if e is None:
        return None
    return e.get("histology")


def set_override(
    store: dict,
    *,
    line: str,
    code: str = "",
    category: str = "",
    comment: str = "",
    by: str = "",
    histology: Optional[bool] = None,
) -> dict:
    """Пишет override по коду (предпочтительно) или по категории. Возвращает store."""
    out = deepcopy(store) if store else empty_store()
    out.setdefault("by_code", {})
    out.setdefault("by_category", {})
    line = str(line or "").strip()
    if not line:
        raise ValueError("Укажите строку ФСН 14")
    code = str(code or "").strip()
    category = str(category or "").strip()
    prev = lookup_entry(out, code=code, category=category) or {}
    hist = histology if histology is not None else prev.get("histology")
    entry = {
        "line": line,
        "comment": _safe_comment(comment),
        "by": str(by or "").strip(),
        "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "histology": hist,
    }
    if code:
        out["by_code"][code] = entry
    elif category:
        out["by_category"][category] = entry
    else:
        raise ValueError("Нужен код или название категории")
    return out


def set_histology(
    store: dict,
    *,
    value: bool,
    code: str = "",
    category: str = "",
    by: str = "",
) -> dict:
    """Правка только морфологии; строку формы не меняет."""
    out = deepcopy(store) if store else empty_store()
    out.setdefault("by_code", {})
    out.setdefault("by_category", {})
    code = str(code or "").strip()
    category = str(category or "").strip()
    prev = lookup_entry(out, code=code, category=category) or {}
    entry = {
        "line": prev.get("line") or "",
        "comment": prev.get("comment") or "",
        "by": str(by or prev.get("by") or "").strip(),
        "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "histology": bool(value),
    }
    if code:
        out["by_code"][code] = entry
    elif category:
        out["by_category"][category] = entry
    else:
        raise ValueError("Нужен код или название категории")
    return out


def clear_override(store: dict, *, code: str = "", category: str = "") -> dict:
    """Отвязать строку формы (вернуться к авто). Морфологию не сбрасывает."""
    out = deepcopy(store) if store else empty_store()
    out.setdefault("by_code", {})
    out.setdefault("by_category", {})
    code = str(code or "").strip()
    category = str(category or "").strip()

    def _clear_line(bucket: dict, key: str) -> None:
        if key not in bucket:
            return
        e = _norm_entry(bucket[key])
        e["line"] = ""
        e["comment"] = ""
        e["at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        if _entry_worth_keeping(e):
            bucket[key] = e
        else:
            del bucket[key]

    if code:
        _clear_line(out["by_code"], code)
    if category:
        _clear_line(out["by_category"], category)
    return out


def merge_store(base: dict, extra: dict) -> dict:
    """Дополняет/перезаписывает base данными из extra."""
    out = deepcopy(base) if base else empty_store()
    out.setdefault("by_code", {})
    out.setdefault("by_category", {})
    for k, v in (extra.get("by_code") or {}).items():
        e = _norm_entry(v)
        if _entry_worth_keeping(e):
            out["by_code"][str(k).strip()] = e
    for k, v in (extra.get("by_category") or {}).items():
        e = _norm_entry(v)
        if _entry_worth_keeping(e):
            out["by_category"][str(k).strip()] = e
    return out


def import_overrides_from_excel(path: Path | str) -> Tuple[dict, int]:
    """
    Читает колонки Строка_ФСН14_ручная / Комментарий / Морфология
    с листов «Все коды» или «Спорные_low_и_21».
    Возвращает (store_fragment, n_rows_imported).
    """
    path = Path(path)
    xl = pd.ExcelFile(path)
    sheets = []
    for name in ("Все коды", "Спорные_low_и_21"):
        if name in xl.sheet_names:
            sheets.append(name)
    if not sheets:
        sheets = list(xl.sheet_names)

    store = empty_store()
    n = 0
    seen_codes: set = set()
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        if df is None or df.empty:
            continue
        cols = {str(c).strip(): c for c in df.columns}
        has_manual = MANUAL_COL in cols
        has_hist = HISTOLOGY_COL in cols
        if not has_manual and not has_hist:
            continue
        code_col = cols.get("Код")
        cat_col = cols.get("Категория") or cols.get("Наименование_КСГ")
        comment_c = cols.get(COMMENT_COL)
        hist_c = cols.get(HISTOLOGY_COL) if has_hist else None
        for _, row in df.iterrows():
            line = _norm_line(row.get(cols[MANUAL_COL])) if has_manual else ""
            hist = _norm_histology(row.get(hist_c)) if hist_c is not None else None
            if not line and hist is None:
                continue
            code = str(row.get(code_col) or "").strip() if code_col is not None else ""
            if code.lower() in ("nan", "none"):
                code = ""
            cat = str(row.get(cat_col) or "").strip() if cat_col is not None else ""
            if cat.lower() in ("nan", "none"):
                cat = ""
            comment = ""
            if comment_c is not None and not pd.isna(row.get(comment_c)):
                comment = _safe_comment(row.get(comment_c))
            if code:
                if code in seen_codes:
                    continue
                seen_codes.add(code)
                if line:
                    store = set_override(
                        store, line=line, code=code, comment=comment, by="excel", histology=hist
                    )
                elif hist is not None:
                    store = set_histology(store, value=hist, code=code, by="excel")
                n += 1
            elif cat:
                if line:
                    store = set_override(
                        store, line=line, category=cat, comment=comment, by="excel", histology=hist
                    )
                elif hist is not None:
                    store = set_histology(store, value=hist, category=cat, by="excel")
                n += 1
    return store, n


def override_line_for(
    store: dict,
    *,
    code: str = "",
    category: str = "",
) -> str:
    hit = lookup_override(store, code=code, category=category)
    return hit[0] if hit else ""


def comment_for(store: dict, *, code: str = "", category: str = "") -> str:
    hit = lookup_override(store, code=code, category=category)
    return (hit[1].get("comment") or "") if hit else ""
