# tests/test_form14_overrides.py
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import yaml

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.dept_config import set_category_histology, sync_histology_overrides_to_config
from analyzers.form14_export import form14_preview_rows_from_ops, write_form14_excel
from analyzers.form14_map import map_code_to_form14
from analyzers.form14_overrides import (
    COMMENT_COL,
    HISTOLOGY_COL,
    MANUAL_COL,
    clear_override,
    empty_store,
    histology_for,
    import_overrides_from_excel,
    load_overrides,
    save_overrides,
    set_histology,
    set_override,
)


def test_override_beats_auto(tmp_path: Path):
    # без override — авто для аппендэктомии 9.2
    auto = map_code_to_form14("A16.18.009", "Аппендэктомия", overrides=None)
    assert auto.line == "9.2"

    store = set_override(empty_store(), line="21", code="A16.18.009", comment="тест")
    m = map_code_to_form14("A16.18.009", "Аппендэктомия", overrides=store)
    assert m.line == "21"
    assert m.rule == "override хирурга"
    assert m.confidence == "high"


def test_lor_without_overrides_unchanged():
    m = map_code_to_form14(
        "A16.08.013.001", "Септопластика", category="Септопластика", overrides={}
    )
    assert m.line == "6"
    assert "калибровка" in m.rule


def test_save_load_roundtrip(tmp_path: Path):
    path = tmp_path / "form14_overrides.yaml"
    store = set_override(empty_store(), line="9.5", code="A16.30.006", comment="лапаротомия")
    save_overrides(store, path)
    loaded = load_overrides(path)
    assert loaded["by_code"]["A16.30.006"]["line"] == "9.5"
    m = map_code_to_form14("A16.30.006", "Лапаротомия", overrides=loaded)
    assert m.line == "9.5"


def test_histology_roundtrip_yaml(tmp_path: Path):
    path = tmp_path / "form14_overrides.yaml"
    store = set_histology(empty_store(), value=True, code="A16.18.009", by="тест")
    save_overrides(store, path)
    loaded = load_overrides(path)
    assert histology_for(loaded, code="A16.18.009") is True
    assert not loaded["by_code"]["A16.18.009"].get("line")

    store2 = set_override(loaded, line="9.2", code="A16.18.009", histology=False)
    save_overrides(store2, path)
    loaded2 = load_overrides(path)
    assert loaded2["by_code"]["A16.18.009"]["line"] == "9.2"
    assert histology_for(loaded2, code="A16.18.009") is False


def test_clear_override_keeps_histology():
    store = set_override(empty_store(), line="21", code="A16.18.009", histology=True)
    store = clear_override(store, code="A16.18.009")
    m = map_code_to_form14("A16.18.009", "Аппендэктомия", overrides=store)
    assert m.line == "9.2"  # снова авто
    assert histology_for(store, code="A16.18.009") is True


def test_import_excel_to_overrides(tmp_path: Path):
    xlsx = tmp_path / "map.xlsx"
    df = pd.DataFrame(
        [
            {
                "Код": "A16.01.004",
                "Наименование_КСГ": "ПХО",
                MANUAL_COL: "17",
                COMMENT_COL: "кожа",
                HISTOLOGY_COL: "да",
            },
            {
                "Код": "A16.18.009",
                "Наименование_КСГ": "Аппендэктомия",
                MANUAL_COL: "",
                COMMENT_COL: "",
                HISTOLOGY_COL: "",
            },
        ]
    )
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Все коды", index=False)

    store, n = import_overrides_from_excel(xlsx)
    assert n == 1
    assert store["by_code"]["A16.01.004"]["line"] == "17"
    assert histology_for(store, code="A16.01.004") is True
    m = map_code_to_form14("A16.01.004", "ПХО", overrides=store)
    assert m.line == "17"
    assert m.rule == "override хирурга"


def test_export_has_manual_and_morph_columns(tmp_path: Path):
    cfg_path = APP / "config.yaml"
    if not cfg_path.exists():
        return
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    store = set_override(
        empty_store(), line="9.3", code="A16.30.001", comment="грыжа", histology=True
    )
    out = tmp_path / "out.xlsx"
    write_form14_excel(out, cfg, overrides=store)
    df = pd.read_excel(out, sheet_name="Все коды")
    assert MANUAL_COL in df.columns
    assert COMMENT_COL in df.columns
    assert HISTOLOGY_COL in df.columns
    hit = df[df["Код"].astype(str) == "A16.30.001"]
    if not hit.empty:
        assert str(hit.iloc[0][MANUAL_COL]) == "9.3"
        assert str(hit.iloc[0][HISTOLOGY_COL]).lower() in ("да", "true", "1")


def test_preview_histology_from_override():
    ops = pd.DataFrame(
        [
            {
                "Код": "A16.18.009",
                "Категория": "Аппендэктомия",
                "Возраст": 40,
                "Гистология": False,
            },
            {
                "Код": "A16.30.001",
                "Категория": "Грыжесечение",
                "Возраст": 55,
                "Гистология": False,
            },
        ]
    )
    store = set_histology(empty_store(), value=True, code="A16.18.009")
    rows = form14_preview_rows_from_ops(
        ops, summary_key="surg1", hide_zeros=True, overrides=store
    )
    by_line = {str(r["line"]): r for r in rows if r.get("line")}
    # аппендэктомия → 9.2; морфология по override
    assert by_line["9.2"]["histology"] == 1
    assert by_line["9.2"]["total"] == 1


def test_set_category_histology_and_sync():
    cfg = {
        "surgery_categories_by_dept": {
            "surg1": [
                {
                    "category": "Аппендэктомия",
                    "codes": ["A16.18.009"],
                    "histology": False,
                    "group": "x",
                    "line": "",
                }
            ]
        }
    }
    assert set_category_histology(
        cfg, "surg1", histology=True, code="A16.18.009"
    )
    assert cfg["surgery_categories_by_dept"]["surg1"][0]["histology"] is True

    cfg2 = {
        "surgery_categories_by_dept": {
            "surg1": [
                {
                    "category": "Аппендэктомия",
                    "codes": ["A16.18.009"],
                    "histology": False,
                    "group": "x",
                    "line": "",
                }
            ]
        }
    }
    store = set_histology(empty_store(), value=True, code="A16.18.009")
    n = sync_histology_overrides_to_config(cfg2, store)
    assert n == 1
    assert cfg2["surgery_categories_by_dept"]["surg1"][0]["histology"] is True


def test_assign_line_preserves_existing_histology():
    store = set_histology(empty_store(), value=True, code="A16.18.009")
    store = set_override(store, line="9.5", code="A16.18.009", comment="x")
    assert histology_for(store, code="A16.18.009") is True
    assert store["by_code"]["A16.18.009"]["line"] == "9.5"


def test_neutralize_excel_formula():
    from analyzers.form14_overrides import neutralize_excel_formula, sanitize_dataframe_for_excel

    assert neutralize_excel_formula("обычный текст") == "обычный текст"
    assert neutralize_excel_formula('=HYPERLINK("https://evil","x")').startswith("'")
    assert neutralize_excel_formula("+cmd|' /C calc'!A0").startswith("'")
    assert neutralize_excel_formula("@SUM(A1)").startswith("'")
    assert neutralize_excel_formula("'-уже") == "'-уже"
    assert neutralize_excel_formula(42) == 42

    store = set_override(
        empty_store(),
        line="21",
        code="A16.01.004",
        comment='=HYPERLINK("https://evil.example","Обновить")',
    )
    assert store["by_code"]["A16.01.004"]["comment"].startswith("'")


def test_import_and_export_neutralize_formulas(tmp_path: Path):
    from openpyxl import Workbook

    xlsx = tmp_path / "map.xlsx"
    payload = '=HYPERLINK("https://evil.example","x")'
    wb = Workbook()
    ws = wb.active
    ws.title = "Все коды"
    ws.append(["Код", "Наименование_КСГ", MANUAL_COL, COMMENT_COL])
    ws.append(["A16.01.004", "ПХО", "17", None])
    # явно строка, иначе openpyxl пишет формулу и pandas читает пустое значение
    cell = ws.cell(row=2, column=4, value=payload)
    cell.data_type = "s"
    wb.save(xlsx)

    store, n = import_overrides_from_excel(xlsx)
    assert n == 1
    assert store["by_code"]["A16.01.004"]["comment"].startswith("'")
    assert "evil.example" in store["by_code"]["A16.01.004"]["comment"]

    cfg_path = APP / "config.yaml"
    if not cfg_path.exists():
        return
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    out = tmp_path / "out.xlsx"
    write_form14_excel(out, cfg, overrides=store)
    exported = pd.read_excel(out, sheet_name="Все коды")
    hit = exported[exported["Код"].astype(str) == "A16.01.004"]
    if not hit.empty:
        comment = str(hit.iloc[0][COMMENT_COL])
        assert comment.startswith("'")
        assert "evil" in comment
