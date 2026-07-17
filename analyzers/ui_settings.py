# analyzers/ui_settings.py
"""Сохранение настроек UI между запусками."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict


def settings_path(app_dir: Path) -> Path:
    return Path(app_dir) / "ui_settings.json"


def load_settings(app_dir: Path) -> Dict[str, Any]:
    path = settings_path(app_dir)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(app_dir: Path, data: Dict[str, Any]) -> None:
    path = settings_path(app_dir)
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass
