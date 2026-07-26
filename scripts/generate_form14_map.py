#!/usr/bin/env python3
"""Черновик сопоставления операций A16 → строки ФСН № 14 (4000/4001)."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.form14_export import DEFAULT_XLSX, load_config, write_form14_excel
from analyzers.form14_overrides import default_path, load_overrides

OUT = APP / DEFAULT_XLSX


def main() -> None:
    cfg = load_config(APP)
    store = load_overrides(default_path(APP))
    write_form14_excel(OUT, cfg, overrides=store)
    n_ov = len(store.get("by_code") or {}) + len(store.get("by_category") or {})
    print(f"→ {OUT}")
    print(f"Overrides: {n_ov} ({default_path(APP).name})")


if __name__ == "__main__":
    main()
