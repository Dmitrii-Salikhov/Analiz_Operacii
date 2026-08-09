# ui_flet/app_context.py
"""Контекст приложения: пути, config, settings — без UI."""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml

from analyzers.app_paths import resolve_app_dir

APP_ROOT = resolve_app_dir(package_file=Path(__file__))
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from analyzers.dept_config import ensure_multi_dept_config
from analyzers.ui_settings import load_settings, save_settings
from analyzers.updater import read_local_version


class AppContext:
    def __init__(self, app_dir: Optional[Path] = None):
        self.app_dir = Path(app_dir or APP_ROOT)
        self.config: dict = {}
        self.settings: dict = {}
        self.reload()

    def reload(self) -> None:
        cfg_path = self.app_dir / "config.yaml"
        if cfg_path.exists():
            self.config = yaml.safe_load(cfg_path.read_text(encoding="utf-8")) or {}
        else:
            self.config = {}
        ensure_multi_dept_config(self.config)
        self.settings = load_settings(self.app_dir / "ui_settings.json")

    def save_settings(self) -> None:
        save_settings(self.settings, self.app_dir / "ui_settings.json")

    @property
    def version(self) -> str:
        return read_local_version(self.app_dir)

    def dept_keys(self) -> List[str]:
        by = self.config.get("surgery_categories_by_dept") or {}
        keys = []
        for k in ("lor", "surg1", "surg2", "pedsurg", "traum"):
            if by.get(k) or (k == "lor" and self.config.get("surgery_categories")):
                keys.append(k)
        return keys

    def dept_full_name(self, key: str) -> str:
        from analyzers.dept_config import dept_full_name

        return dept_full_name(self.config, key)

    def dept_dropdown_options(self) -> List[tuple]:
        """(key, полное название) для фильтров конструктора."""
        return [(k, self.dept_full_name(k)) for k in self.dept_keys()]

    def current_summary_key(self) -> str:
        from analyzers.dept_config import dept_summary_key

        dept = self.settings.get("department") or (self.config.get("departments") or {}).get("main")
        return dept_summary_key(self.config, dept)
