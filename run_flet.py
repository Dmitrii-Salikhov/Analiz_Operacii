#!/usr/bin/env python3
"""Запуск Flet UI: python3 run_flet.py"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def _bind_bundled_flet_view() -> None:
    """В frozen-сборке окно рисует flet.exe — берём наш патченный клиент, не кэш ~/.flet."""
    if not getattr(sys, "frozen", False):
        return
    if os.environ.get("FLET_VIEW_PATH"):
        return
    root = Path(sys.executable).resolve().parent
    candidates = [
        root / "flet_view" / "flet",
        root / "_internal" / "flet_view" / "flet",
    ]
    for view in candidates:
        if (view / "flet.exe").is_file():
            os.environ["FLET_VIEW_PATH"] = str(view)
            return


_bind_bundled_flet_view()

from ui_flet.app import main  # noqa: E402

if __name__ == "__main__":
    main()
