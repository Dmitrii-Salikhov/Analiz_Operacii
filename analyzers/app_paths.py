# analyzers/app_paths.py
"""Корень приложения: рядом с exe (frozen) или корень репозитория."""
from __future__ import annotations

import sys
from pathlib import Path


def resolve_app_dir(*, package_file: Path | None = None) -> Path:
    """
    package_file — __file__ модуля внутри пакета (ui_flet/..., analyzers/...).
    Для обычного запуска: parents[1] от package_file → корень проекта.
    Для PyInstaller onedir: каталог с AnalizOperacii.exe.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    if package_file is not None:
        p = Path(package_file).resolve()
        # ui_flet/x.py → parents[1]; analyzers/x.py → parents[1]
        if p.parent.name in ("ui_flet", "analyzers", "screens"):
            if p.parent.name == "screens":
                return p.parents[2]
            return p.parents[1]
        return p.parent
    return Path(__file__).resolve().parents[1]
