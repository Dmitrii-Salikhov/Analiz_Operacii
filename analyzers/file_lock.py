# analyzers/file_lock.py
"""Проверка, что Excel-файл не занят другим процессом."""
from __future__ import annotations

import os
from pathlib import Path


def excel_file_locked(path: str) -> bool:
    """True, если файл, скорее всего, открыт в Excel (нельзя перезаписать)."""
    p = Path(path)
    if not p.exists():
        return False
    # Excel на Windows/macOS часто держит ~$имя.xlsx
    lock_sibling = p.with_name(f"~${p.name}")
    if lock_sibling.exists():
        return True
    try:
        with open(p, "a+b"):
            pass
        return False
    except OSError:
        return True


def is_writable(path: str) -> bool:
    p = Path(path)
    if not p.exists():
        return True
    return os.access(p, os.W_OK) and not excel_file_locked(path)
