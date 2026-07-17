# analyzers/year_template.py
"""Создание сводной на новый год из шаблона предыдущего."""
from __future__ import annotations

import calendar
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Optional

import openpyxl

from analyzers.summary_writer import MONTH_RU


def create_year_summary(
    template_path: str,
    new_year: int,
    output_path: Optional[str] = None,
    sheet_names: Optional[Dict[int, str]] = None,
    clear_values: bool = True,
) -> Path:
    """
    Копирует шаблон сводной и готовит файл на new_year:
    - C2 каждого месячного листа = 1-е число месяца new_year (даты недель пересчитает Excel);
    - при clear_values очищает числовые ячейки категорий C–G (формулы H/ИТОГ/форма не трогаем жёстко —
      очищаем только строки категорий и детей/человек).
    """
    src = Path(template_path)
    if not src.exists():
        raise FileNotFoundError(src)
    out = Path(output_path) if output_path else src.with_name(f"Операции сводная {new_year}.xlsx")
    shutil.copy2(src, out)

    names = sheet_names or {
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
    }

    wb = openpyxl.load_workbook(out)
    for month, sheet in names.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        # дата начала месяца — Excel serial через datetime
        ws.cell(2, 3).value = datetime(new_year, month, 1)
        # заголовок
        title = ws.cell(1, 1).value
        if isinstance(title, str):
            for y in range(new_year - 5, new_year + 2):
                title = title.replace(str(y), str(new_year))
            # «за апрель 2026» → новый год
            for mname in MONTH_RU.values():
                if mname in title.lower():
                    pass
            ws.cell(1, 1).value = title

        if clear_values:
            for row in range(4, 38):
                for col in range(3, 8):
                    cell = ws.cell(row, col)
                    # не затираем формулы
                    if isinstance(cell.value, str) and str(cell.value).startswith("="):
                        continue
                    cell.value = None
            for row in (43, 45):
                for col in range(3, 8):
                    cell = ws.cell(row, col)
                    if isinstance(cell.value, str) and str(cell.value).startswith("="):
                        continue
                    cell.value = None
            # листовые числа формы 4001 (N/O/P/Q/R строк 12–17)
            for row in range(12, 18):
                for col in range(14, 19):
                    cell = ws.cell(row, col)
                    if isinstance(cell.value, str) and str(cell.value).startswith("="):
                        continue
                    cell.value = None

    wb.save(out)
    return out


def suggest_summary_path(app_dir: Path, year: int) -> Path:
    return app_dir / f"Операции сводная {year}.xlsx"
