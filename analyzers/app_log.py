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

    def read_lines(self, *, trim: bool = False) -> List[str]:
        if not self.path.exists():
            return []
        try:
            text = self.path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            return []
        lines = [ln for ln in text.splitlines() if ln.strip() != ""]
        if len(lines) > self.max_lines:
            lines = lines[-self.max_lines :]
            if trim:
                self._write(lines)
        return lines

    def trim(self) -> int:
        """Оставляет последние max_lines, удаляет старые. Возвращает число оставшихся строк."""
        lines = self.read_lines(trim=False)
        if len(lines) > self.max_lines:
            lines = lines[-self.max_lines :]
            self._write(lines)
        elif lines:
            # перезапись не нужна, но файл мог содержать пустые хвосты — нормализуем при переполнении только
            pass
        return len(lines)

    def append(self, message: str, level: str = "INFO") -> str:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{ts} - {level} - {message}"
        lines = self.read_lines(trim=False)
        lines.append(line)
        if len(lines) > self.max_lines:
            lines = lines[-self.max_lines :]
        self._write(lines)
        return line

    def _write(self, lines: List[str]) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(("\n".join(lines) + "\n") if lines else "", encoding="utf-8")
        except OSError:
            pass

    def clear(self) -> None:
        self._write([])
