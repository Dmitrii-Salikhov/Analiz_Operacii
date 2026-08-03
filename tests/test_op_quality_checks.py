# tests/test_op_quality_checks.py
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.op_quality_checks import (
    find_long_operations,
    find_missing_or_table,
    long_op_hours_from_config,
)
from analyzers.surgery import duration_hours, parse_op_datetime


def test_parse_op_datetime_and_duration():
    start = parse_op_datetime("30.04.2026", "08:00")
    end = parse_op_datetime("30.04.2026", "12:30")
    assert start is not None and end is not None
    h = duration_hours(start, end)
    assert h is not None
    assert abs(h - 4.5) < 1e-6


def test_long_ops_threshold_exclusive():
    ops = pd.DataFrame(
        [
            {
                "КВС": "26/111",
                "Пациент": "Иванов И.И.",
                "Хирург": "Петров",
                "Услуга": "Операция А",
                "Длительность_ч": 4.0,
                "Опер.стол": "1 стол",
                "Начало": "a",
            },
            {
                "КВС": "26/222",
                "Пациент": "Сидоров",
                "Хирург": "Петров",
                "Услуга": "Операция Б",
                "Длительность_ч": 4.01,
                "Опер.стол": "1 стол",
                "Начало": "b",
            },
        ]
    )
    rows = find_long_operations(ops, max_hours=4)
    assert len(rows) == 1
    assert rows[0]["КВС"] == "26/222"
    assert rows[0]["Пациент"] == "Сидоров"
    assert rows[0]["Хирург"] == "Петров"
    assert "Операция Б" in rows[0]["Услуга"]


def test_missing_or_table():
    ops = pd.DataFrame(
        [
            {
                "КВС": "26/1",
                "Пациент": "A",
                "Хирург": "X",
                "Услуга": "Услуга 1",
                "Опер.стол": "5 Опер.стол",
                "Дата": "2026-04-01",
            },
            {
                "КВС": "26/2",
                "Пациент": "B",
                "Хирург": "Y",
                "Услуга": "Услуга 2",
                "Опер.стол": "   ",
                "Дата": "2026-04-01",
            },
            {
                "КВС": "26/3",
                "Пациент": "C",
                "Хирург": "Z",
                "Услуга": "Услуга 3",
                "Опер.стол": "",
                "Дата": "2026-04-02",
            },
        ]
    )
    rows = find_missing_or_table(ops)
    assert len(rows) == 2
    kvss = {r["КВС"] for r in rows}
    assert kvss == {"26/2", "26/3"}
    assert all(r["Причина"] == "не занесена на опер.стол" for r in rows)
    assert all(r["Пациент"] and r["Хирург"] and r["Услуга"] for r in rows)


def test_long_op_hours_from_config():
    assert long_op_hours_from_config({}) == 4.0
    assert long_op_hours_from_config({"thresholds": {"long_op_hours": 6}}) == 6.0
