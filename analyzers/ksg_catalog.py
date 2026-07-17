# analyzers/ksg_catalog.py
"""Справочник услуг/КСГ из KSGoperacii.csv."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

APP_DIR = Path(__file__).resolve().parent.parent
DEFAULT_PATH = APP_DIR / "KSGoperacii.csv"


class KsgCatalog:
    def __init__(self, path: Optional[Path] = None):
        self.path = Path(path) if path else DEFAULT_PATH
        self.by_code: Dict[str, dict] = {}
        self._load()

    def _load(self):
        if not self.path.exists():
            return
        df = pd.read_csv(self.path, sep=";", encoding="utf-8", dtype=str)
        df = df.fillna("")
        code_col = "Код услуги"
        name_col = "Наименование услуги"
        ksg_cols = [c for c in df.columns if str(c).startswith("КСГ")]
        for _, row in df.iterrows():
            code = str(row.get(code_col, "")).strip()
            if not code:
                continue
            ksgs = [str(row[c]).strip() for c in ksg_cols if str(row.get(c, "")).strip()]
            self.by_code[code] = {
                "code": code,
                "name": str(row.get(name_col, "")).strip(),
                "ksg": ksgs,
                "used": str(row.get("Использовано в КСГ", "")).strip().upper() == "TRUE",
            }

    def lookup(self, code: str) -> Optional[dict]:
        if not code:
            return None
        return self.by_code.get(str(code).strip())

    def name_for(self, code: str) -> str:
        info = self.lookup(code)
        return info["name"] if info else ""

    def ksg_for(self, code: str) -> str:
        info = self.lookup(code)
        if not info or not info["ksg"]:
            return ""
        return ", ".join(info["ksg"])

    def hint_for(self, code: str) -> str:
        """Краткая подсказка для неклассифицированного кода."""
        info = self.lookup(code)
        if not info:
            return "нет в KSGoperacii.csv"
        parts = [info["name"]] if info["name"] else []
        if info["ksg"]:
            parts.append("КСГ: " + ", ".join(info["ksg"]))
        return " | ".join(parts) if parts else code

    def suggest_config_line(self, code: str) -> str:
        info = self.lookup(code)
        name = info["name"] if info else ""
        return f'  # TODO: "{name}" → category\n  # codes: ["{code}"]'


_catalog: Optional[KsgCatalog] = None


def get_catalog(path: Optional[Path] = None) -> KsgCatalog:
    global _catalog
    if _catalog is None or (path and Path(path) != _catalog.path):
        _catalog = KsgCatalog(path)
    return _catalog
