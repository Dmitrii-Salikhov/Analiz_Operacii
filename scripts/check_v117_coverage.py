#!/usr/bin/env python3
"""Проверка покрытия >=90% для кода v1.1.7 (разделители Excel, insert, uncl)."""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# (путь от корня репозитория, диапазоны строк inclusive)
# Регионы: разделители Excel, insert категории, сдвиг высот строк, uncl/session
V117_REGIONS: list[tuple[str, list[tuple[int, int]]]] = [
    (
        "analyzers/summary_layout.py",
        [
            (99, 215),   # shift_row_dimensions / fix_patients_row_height / sheet_insert|delete
            (382, 529),  # ensure_one_blank_before_totals / between_labels
            (848, 976),  # add_category_row_to_summary
        ],
    ),
    ("analyzers/summary_writer.py", [(182, 222)]),
    ("ui_flet/session.py", [(701, 806), (843, 885)]),
]


def _load_json_report(data_file: str) -> dict:
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
        out_path = tmp.name
    try:
        subprocess.run(
            [
                sys.executable,
                "-m",
                "coverage",
                "json",
                "-o",
                out_path,
                "--data-file",
                data_file,
            ],
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )
        with open(out_path, encoding="utf-8") as fh:
            return json.load(fh)
    finally:
        Path(out_path).unlink(missing_ok=True)


def _region_stats(
    executed: set[int],
    missing: set[int],
    start: int,
    end: int,
) -> tuple[int, int, list[int]]:
    in_region_exec = {ln for ln in executed if start <= ln <= end}
    in_region_miss = sorted(ln for ln in missing if start <= ln <= end)
    in_region_all = in_region_exec | set(in_region_miss)
    total = len(in_region_all)
    covered = total - len(in_region_miss)
    return covered, total, in_region_miss


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fail-under", type=float, default=90.0)
    parser.add_argument("--data-file", default=".coverage")
    args = parser.parse_args()

    data_path = ROOT / args.data_file
    if not data_path.exists():
        print(f"ERROR: coverage data not found: {data_path}")
        return 1

    report = _load_json_report(str(data_path))
    files = report.get("files") or {}

    grand_covered = 0
    grand_total = 0
    failures: list[str] = []

    print("v1.1.7 coverage regions:")
    for rel_path, ranges in V117_REGIONS:
        file_key = None
        for key in files:
            if key.replace("\\", "/").endswith(rel_path):
                file_key = key
                break
        if not file_key:
            failures.append(f"{rel_path}: not measured")
            print(f"  {rel_path}: NOT MEASURED")
            continue

        info = files[file_key]
        executed = set(info.get("executed_lines") or [])
        missing = set(info.get("missing_lines") or [])

        file_covered = 0
        file_total = 0
        for start, end in ranges:
            cov, tot, miss = _region_stats(executed, missing, start, end)
            file_covered += cov
            file_total += tot
            pct = (100.0 * cov / tot) if tot else 100.0
            print(f"  {rel_path}:{start}-{end}  {cov}/{tot}  ({pct:.1f}%)")
            if miss:
                print(f"    missing: {miss[:20]}{'…' if len(miss) > 20 else ''}")
        file_pct = (100.0 * file_covered / file_total) if file_total else 100.0
        print(f"  => {rel_path}: {file_covered}/{file_total} ({file_pct:.1f}%)")
        if file_total and file_pct < args.fail_under:
            failures.append(f"{rel_path}: {file_pct:.1f}% < {args.fail_under}%")
        grand_covered += file_covered
        grand_total += file_total

    overall = (100.0 * grand_covered / grand_total) if grand_total else 100.0
    print(f"TOTAL v1.1.7: {grand_covered}/{grand_total} ({overall:.1f}%)")

    if overall < args.fail_under:
        failures.append(f"TOTAL: {overall:.1f}% < {args.fail_under}%")

    if failures:
        print("\nFAIL:")
        for msg in failures:
            print(f"  - {msg}")
        return 1

    print(f"\nOK: v1.1.7 coverage >= {args.fail_under}%")
    return 0


if __name__ == "__main__":
    sys.exit(main())
