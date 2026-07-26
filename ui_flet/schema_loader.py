# ui_flet/schema_loader.py
"""Загрузка YAML-схем и реестра экранов."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml

APP_ROOT = Path(__file__).resolve().parents[1]
SCHEMAS = APP_ROOT / "schemas"


def load_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def load_registry() -> dict:
    return load_yaml(SCHEMAS / "app_registry.yaml")


def load_schema(name: str) -> dict:
    p = SCHEMAS / name
    if not p.exists():
        p = SCHEMAS / f"{name}.schema.yaml"
    if not p.exists() and not name.endswith(".yaml"):
        p = SCHEMAS / f"{name}.yaml"
    return load_yaml(p)


def screens_by_status(registry: Optional[dict] = None) -> Dict[str, List[dict]]:
    reg = registry or load_registry()
    out: Dict[str, List[dict]] = {"done": [], "shell": [], "backend": []}
    for s in reg.get("screens") or []:
        st = str(s.get("status") or "shell")
        out.setdefault(st, []).append(s)
    return out


def actions_checklist(registry: Optional[dict] = None) -> List[dict]:
    reg = registry or load_registry()
    return list(reg.get("actions") or [])


def param_defs(schema: dict) -> List[dict]:
    return list(schema.get("parameters") or [])


def filter_defs(schema: dict) -> List[dict]:
    return list(schema.get("filters") or [])
