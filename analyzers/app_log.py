"""Файловый журнал с ограничением длины (ротация хвоста)."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import List

LOG_MAX_LINES = 500


class AppLog:
    def __init__(self, path: Path, max_lines: int = LOG_MAX_LINES):
        self.path = Path(path)
        self.max_lines = max_lines

    def read_lines(self) -> List[str]:
        if not self.path.exists():
            return []
        try:
            text = self.path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            return []
        return [ln for ln in text.splitlines() if ln.strip() != ""]

    def append(self, message: str, level: str = "INFO") -> str:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{ts} - {level} - {message}"
        lines = self.read_lines()
        lines.append(line)
        if len(lines) > self.max_lines:
            lines = lines[-self.max_lines :]
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        except OSError:
            pass
        return line

    def clear(self) -> None:
        try:
            self.path.write_text("", encoding="utf-8")
        except OSError:
            pass
