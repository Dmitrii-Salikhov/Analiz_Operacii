# analyzers/category_registry.py
"""Регистрация новой категории операции в config.yaml."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import yaml

from analyzers.dept_config import (
    ensure_multi_dept_config,
    form_4001_enabled,
    get_summary_cfg,
    get_surgery_categories,
    set_summary_cfg,
    set_surgery_categories,
)
from analyzers.form_4001 import (
    ENDO_CATS,
    FORM_LINE_LABELS,
    LINE_HIST_CATS,
    LINE_TOTAL_CATS,
)

FORM_LINES = ("5.1", "5.2", "6", "6.1", "17")


@dataclass
class CategorySpec:
    name: str
    codes: List[str] = field(default_factory=list)
    name_keywords: List[str] = field(default_factory=list)
    kind: str = "plan"  # plan | emergency
    form_line: str = "6"
    histology: bool = False
    endoscopic: bool = False
    group: str = ""
    anchor_category: str = ""


@dataclass
class RegisterResult:
    name: str
    excel_row: int
    config_path: Path
    warnings: List[str] = field(default_factory=list)


class CategoryRegistryError(ValueError):
    pass


def _norm_codes(codes: Sequence[str]) -> List[str]:
    out: List[str] = []
    for c in codes:
        s = str(c or "").strip()
        if s and s not in out:
            out.append(s)
    return out


def _norm_keywords(words: Sequence[str]) -> List[str]:
    out: List[str] = []
    for w in words:
        s = str(w or "").strip().lower()
        if s and s not in out:
            out.append(s)
    return out


def suggest_keywords_from_name(name: str) -> List[str]:
    raw = (
        str(name or "")
        .replace(",", " ")
        .replace("(", " ")
        .replace(")", " ")
        .replace("«", " ")
        .replace("»", " ")
    )
    return _norm_keywords([w for w in raw.split() if len(w) > 3])[:6]


def _summary_section(config: dict, summary_key: str = "lor") -> dict:
    ensure_multi_dept_config(config)
    summaries = config.setdefault("summaries", {})
    if summary_key not in summaries:
        summaries[summary_key] = deepcopy(get_summary_cfg(config, summary_key=summary_key) or {})
    section = summaries[summary_key]
    if summary_key == "lor":
        config["summary"] = section
    return section


def _surgery_section(config: dict, summary_key: str = "lor") -> List[dict]:
    ensure_multi_dept_config(config)
    cats = get_surgery_categories(config, summary_key=summary_key)
    by_dept = config.setdefault("surgery_categories_by_dept", {})
    if summary_key not in by_dept:
        by_dept[summary_key] = list(cats)
    if summary_key == "lor":
        config["surgery_categories"] = by_dept[summary_key]
    return by_dept[summary_key]


def existing_category_names(config: dict, summary_key: str = "lor") -> List[str]:
    names: List[str] = []
    rows = (_summary_section(config, summary_key).get("category_rows") or {})
    names.extend(str(k) for k in rows.keys())
    for cat in _surgery_section(config, summary_key):
        n = str(cat.get("category") or "")
        if n and n not in names:
            names.append(n)
    return names


def _resolve_category_key(rows: dict, name: str) -> str:
    """Точное имя или совпадение без крайних пробелов (как у «Биопсия гортани »)."""
    if not name:
        return ""
    if name in rows:
        return name
    needle = str(name).strip()
    for key in rows:
        if str(key).strip() == needle:
            return str(key)
    return ""


def allocate_excel_row(config: dict, anchor_category: str, summary_key: str = "lor") -> int:
    """
    Строка для физической вставки: сразу после якоря (anchor_row + 1).
    Существующие category_rows / totals_rows >= этой строки сдвигаются отдельно.
    """
    summary = _summary_section(config, summary_key)
    rows_map: Dict[str, int] = {
        str(k): int(v) for k, v in (summary.get("category_rows") or {}).items()
    }
    if not rows_map:
        raise CategoryRegistryError("В config нет category_rows для отделения")

    key = _resolve_category_key(rows_map, anchor_category)
    if key:
        anchor_row = int(rows_map[key])
    else:
        anchor_row = max(rows_map.values())

    return int(anchor_row) + 1


def shift_rows_for_insert(config: dict, insert_at: int, summary_key: str = "lor") -> None:
    """Сдвигает category_rows и totals_rows на +1, если row >= insert_at."""
    summary = _summary_section(config, summary_key)
    rows = summary.setdefault("category_rows", {})
    for name, row in list(rows.items()):
        if int(row) >= insert_at:
            rows[name] = int(row) + 1
    totals = summary.setdefault("totals_rows", {})
    for key, row in list(totals.items()):
        try:
            if int(row) >= insert_at:
                totals[key] = int(row) + 1
        except (TypeError, ValueError):
            pass
    set_summary_cfg(config, summary_key, summary)


def default_anchor_category(config: dict, summary_key: str = "lor") -> str:
    rows = (_summary_section(config, summary_key).get("category_rows") or {})
    if not rows:
        return ""
    # предпочитаем биопсию, иначе последнюю по номеру строки
    for key in rows:
        if "биопсия" in str(key).lower():
            return str(key)
    return max(rows.items(), key=lambda kv: int(kv[1]))[0]


def _seed_list(cfg_list: Optional[list], defaults: Sequence[str]) -> List[str]:
    if cfg_list is None:
        return list(defaults)
    out = [str(x) for x in cfg_list]
    for d in defaults:
        if d not in out:
            out.append(d)
    return out


def _ensure_line_categories(summary: dict) -> dict:
    form = summary.setdefault("form_4001", {})
    lc = form.setdefault("line_categories", {})
    for line, cats in LINE_TOTAL_CATS.items():
        if line not in lc:
            lc[line] = list(cats)
        else:
            lc[line] = _seed_list(lc[line], cats)
    hc = form.setdefault("hist_categories", {})
    for line, cats in LINE_HIST_CATS.items():
        if line not in hc:
            hc[line] = list(cats)
        else:
            hc[line] = _seed_list(hc[line], cats)
    if "endo_categories" not in form:
        form["endo_categories"] = list(ENDO_CATS)
    else:
        form["endo_categories"] = _seed_list(form.get("endo_categories"), ENDO_CATS)
    return form


def apply_category_to_config(
    config: dict, spec: CategorySpec, summary_key: str = "lor"
) -> RegisterResult:
    """Мутирует config in-place. Возвращает результат с выбранным excel_row."""
    name = str(spec.name or "").strip()
    if not name:
        raise CategoryRegistryError("Укажите название категории")
    if spec.form_line not in FORM_LINES:
        raise CategoryRegistryError(f"Строка формы 4001 должна быть одной из: {', '.join(FORM_LINES)}")
    if spec.kind not in ("plan", "emergency"):
        raise CategoryRegistryError("Тип: plan или emergency")

    existing = set(existing_category_names(config, summary_key))
    if name in existing:
        raise CategoryRegistryError(f"Категория уже есть: «{name}»")

    codes = _norm_codes(spec.codes)
    keywords = _norm_keywords(spec.name_keywords) or suggest_keywords_from_name(name)
    group = (spec.group or "").strip() or FORM_LINE_LABELS.get(spec.form_line, "прочее")

    anchor_raw = (spec.anchor_category or "") or default_anchor_category(config, summary_key)
    summary = _summary_section(config, summary_key)
    rows_preview = summary.get("category_rows") or {}
    anchor_key = _resolve_category_key(rows_preview, anchor_raw) or default_anchor_category(config, summary_key)
    excel_row = allocate_excel_row(config, anchor_key, summary_key)
    warnings: List[str] = []
    if anchor_raw and not _resolve_category_key(rows_preview, anchor_raw):
        warnings.append(f"Якорь «{anchor_raw}» не найден — вставлено после последней категории")

    shift_rows_for_insert(config, excel_row, summary_key)
    summary = _summary_section(config, summary_key)
    rows = summary.setdefault("category_rows", {})
    rows[name] = excel_row

    plan = summary.setdefault("plan_categories", [])
    emerg = summary.setdefault("emergency_categories", [])
    if spec.kind == "emergency":
        if name not in emerg:
            emerg.append(name)
        if name in plan:
            plan.remove(name)
    else:
        if name not in plan:
            plan.append(name)
        if name in emerg:
            emerg.remove(name)

    if form_4001_enabled(summary):
        form = _ensure_line_categories(summary)
        line_cats: List[str] = form["line_categories"].setdefault(
            spec.form_line, list(LINE_TOTAL_CATS.get(spec.form_line, []))
        )
        if name not in line_cats:
            line_cats.append(name)
        if spec.histology:
            hist = form["hist_categories"].setdefault(
                spec.form_line, list(LINE_HIST_CATS.get(spec.form_line, []))
            )
            if name not in hist:
                hist.append(name)
        if spec.endoscopic:
            endo = form.setdefault("endo_categories", list(ENDO_CATS))
            if name not in endo:
                endo.append(name)

    surgery = _surgery_section(config, summary_key)
    surgery.append(
        {
            "category": name,
            "codes": codes,
            "group": group,
            "line": spec.form_line,
            "histology": bool(spec.histology),
            "name_keywords": keywords,
        }
    )
    set_summary_cfg(config, summary_key, summary)
    set_surgery_categories(config, summary_key, surgery)

    return RegisterResult(name=name, excel_row=excel_row, config_path=Path(), warnings=warnings)


def update_category_keywords(
    config: dict,
    category_name: str,
    keywords: Sequence[str],
    summary_key: str = "lor",
) -> List[str]:
    """Обновляет name_keywords у категории. Возвращает нормализованный список."""
    name = str(category_name or "").strip()
    if not name:
        raise CategoryRegistryError("Укажите название категории")
    surgery = _surgery_section(config, summary_key)
    target = None
    for cat in surgery:
        if str(cat.get("category") or "").strip() == name:
            target = cat
            break
        emerg = str(cat.get("emergency_category") or "").strip()
        if emerg and emerg == name:
            target = cat
            break
    if target is None:
        raise CategoryRegistryError(f"Категория не найдена: «{name}»")
    normed = _norm_keywords(keywords)
    target["name_keywords"] = normed
    set_surgery_categories(config, summary_key, surgery)
    return normed


def update_category_keywords_file(
    config_path: Path,
    category_name: str,
    keywords: Sequence[str],
    *,
    config: Optional[dict] = None,
    summary_key: str = "lor",
) -> tuple[dict, List[str]]:
    path = Path(config_path)
    cfg = deepcopy(config) if config is not None else load_config(path)
    normed = update_category_keywords(cfg, category_name, keywords, summary_key=summary_key)
    save_config(cfg, path)
    return cfg, normed


def save_config(config: dict, path: Path) -> None:
    path = Path(path)
    text = yaml.safe_dump(
        config,
        allow_unicode=True,
        default_flow_style=False,
        sort_keys=False,
        width=120,
    )
    path.write_text(text, encoding="utf-8")


def load_config(path: Path) -> dict:
    path = Path(path)
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise CategoryRegistryError(f"Некорректный config: {path}")
    return data


@dataclass
class RemoveResult:
    name: str
    excel_row: int
    config_path: Path
    warnings: List[str] = field(default_factory=list)


def shift_rows_for_delete(config: dict, delete_at: int, summary_key: str = "lor") -> None:
    """Сдвигает category_rows / totals_rows на −1 для строк > delete_at."""
    summary = _summary_section(config, summary_key)
    rows = summary.setdefault("category_rows", {})
    for name, row in list(rows.items()):
        r = int(row)
        if r > delete_at:
            rows[name] = r - 1
    totals = summary.setdefault("totals_rows", {})
    for key, row in list(totals.items()):
        try:
            r = int(row)
            if r > delete_at:
                totals[key] = r - 1
        except (TypeError, ValueError):
            pass
    set_summary_cfg(config, summary_key, summary)


def shift_totals_rows_by_delta(config: dict, delta: int, summary_key: str = "lor") -> None:
    """Сдвигает только totals_rows (пустая строка-разделитель перед итогами)."""
    if not delta:
        return
    summary = _summary_section(config, summary_key)
    totals = summary.setdefault("totals_rows", {})
    for key, row in list(totals.items()):
        try:
            totals[key] = int(row) + int(delta)
        except (TypeError, ValueError):
            pass
    set_summary_cfg(config, summary_key, summary)


def _remove_name_from_lists(summary: dict, name: str) -> None:
    for key in ("plan_categories", "emergency_categories"):
        lst = summary.get(key) or []
        summary[key] = [x for x in lst if str(x) != name]
    form = summary.get("form_4001") or {}
    for map_key in ("line_categories", "hist_categories"):
        mapping = form.get(map_key) or {}
        for line, cats in list(mapping.items()):
            mapping[line] = [c for c in (cats or []) if str(c) != name]
    endo = form.get("endo_categories") or []
    if endo:
        form["endo_categories"] = [c for c in endo if str(c) != name]


def remove_category_from_config(config: dict, name: str, summary_key: str = "lor") -> RemoveResult:
    """Удаляет категорию из config и сдвигает номера строк."""
    name = str(name or "").strip()
    if not name:
        raise CategoryRegistryError("Укажите название категории")
    summary = _summary_section(config, summary_key)
    rows = summary.setdefault("category_rows", {})
    key = _resolve_category_key(rows, name)
    if not key:
        raise CategoryRegistryError(f"Категория не найдена: «{name}»")
    excel_row = int(rows[key])
    del rows[key]
    _remove_name_from_lists(summary, key)
    surgery = _surgery_section(config, summary_key)
    filtered = [c for c in surgery if str(c.get("category") or "") != key]
    set_surgery_categories(config, summary_key, filtered)
    shift_rows_for_delete(config, excel_row, summary_key)
    set_summary_cfg(config, summary_key, summary)
    return RemoveResult(name=key, excel_row=excel_row, config_path=Path())


def unregister_category(
    config_path: Path,
    name: str,
    *,
    config: Optional[dict] = None,
    summary_key: str = "lor",
) -> tuple[dict, RemoveResult]:
    path = Path(config_path)
    cfg = deepcopy(config) if config is not None else load_config(path)
    result = remove_category_from_config(cfg, name, summary_key=summary_key)
    save_config(cfg, path)
    result.config_path = path
    return cfg, result


def register_category(
    config_path: Path,
    spec: CategorySpec,
    *,
    config: Optional[dict] = None,
    summary_key: str = "lor",
) -> tuple[dict, RegisterResult]:
    """
    Загружает config (или использует переданный), регистрирует категорию, сохраняет YAML.
    Возвращает (обновлённый config, result).
    """
    path = Path(config_path)
    cfg = deepcopy(config) if config is not None else load_config(path)
    result = apply_category_to_config(cfg, spec, summary_key=summary_key)
    save_config(cfg, path)
    result.config_path = path
    return cfg, result


def category_spec_from_dict(d: Dict[str, Any]) -> CategorySpec:
    codes = d.get("codes") or []
    if isinstance(codes, str):
        codes = [c.strip() for c in codes.replace(";", ",").split(",") if c.strip()]
    kws = d.get("name_keywords") or []
    if isinstance(kws, str):
        kws = [c.strip() for c in kws.replace(";", ",").split(",") if c.strip()]
    return CategorySpec(
        name=str(d.get("name") or ""),
        codes=list(codes),
        name_keywords=list(kws),
        kind=str(d.get("kind") or "plan"),
        form_line=str(d.get("form_line") or "6"),
        histology=bool(d.get("histology", False)),
        endoscopic=bool(d.get("endoscopic", False)),
        group=str(d.get("group") or ""),
        anchor_category=str(d.get("anchor_category") or ""),
    )
