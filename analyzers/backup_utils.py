# analyzers/backup_utils.py
"""Бэкапы сводной: список, ротация, восстановление."""
from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import List, Optional

# Имя: «Операции сводная 2026.20260717_203001.bak.xlsx»
_BAK_RE = re.compile(r"^.+\.\d{8}_\d{6}\.bak\.xlsx$", re.IGNORECASE)


def list_backups(summary_path: str | Path, limit: int = 50) -> List[Path]:
    p = Path(summary_path)
    if not p.parent.exists():
        return []
    stem = p.stem  # без .xlsx
    found: List[Path] = []
    for f in p.parent.iterdir():
        if not f.is_file():
            continue
        name = f.name
        if not name.lower().endswith(".bak.xlsx"):
            continue
        # наш формат: {stem}.{stamp}.bak.xlsx
        if name.startswith(stem + ".") and _BAK_RE.match(name):
            found.append(f)
        elif stem in name and ".bak." in name.lower():
            found.append(f)
    found.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return found[:limit]


def rotate_backups(summary_path: str | Path, keep: int = 10) -> List[Path]:
    """Удаляет старые .bak, оставляет keep самых новых. Возвращает удалённые."""
    keep = max(1, int(keep))
    all_baks = list_backups(summary_path, limit=500)
    removed: List[Path] = []
    for old in all_baks[keep:]:
        try:
            old.unlink()
            removed.append(old)
        except OSError:
            pass
    return removed


def restore_backup(backup_path: str | Path, summary_path: str | Path) -> Path:
    """Копирует бэкап поверх сводной (сначала бэкап текущего файла)."""
    bak = Path(backup_path)
    dest = Path(summary_path)
    if not bak.exists():
        raise FileNotFoundError(bak)
    if dest.exists():
        from datetime import datetime

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safety = dest.with_name(f"{dest.stem}.before_restore.{stamp}.bak{dest.suffix}")
        shutil.copy2(dest, safety)
    shutil.copy2(bak, dest)
    return dest
