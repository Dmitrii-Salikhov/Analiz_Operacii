import pandas as pd

from analyzers.problem_codes import build_problem_codes_table, format_config_draft


def test_build_problem_codes_empty():
    df = build_problem_codes_table(pd.DataFrame())
    assert list(df.columns) == ["Код", "Операций", "Услуга_пример", "Подсказка_КСГ", "Черновик_config"]
    assert df.empty


def test_build_problem_codes_and_draft():
    ops = pd.DataFrame(
        {
            "Категория": ["Не классифицировано", "Не классифицировано", "Другое"],
            "Код": ["X99.99.001", "X99.99.001", "A01"],
            "Услуга": ["операция А", "операция Б", "ок"],
        }
    )
    table = build_problem_codes_table(ops)
    assert len(table) == 1
    assert int(table.iloc[0]["Операций"]) == 2
    assert "X99.99.001" in str(table.iloc[0]["Черновик_config"])
    draft = format_config_draft(table)
    assert "surgery_categories" in draft
    assert "X99.99.001" in draft
