#!/usr/bin/env python3
"""Запуск Flet UI: python3 run_flet.py"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from ui_flet.app import main

if __name__ == "__main__":
    main()
