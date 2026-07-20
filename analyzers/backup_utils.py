# analyzers/backup_utils.py
"""Бэкапы сводной: отдельная папка backups/, список, ротация, восстановление."""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Optional

# Имя: «Операции сводная 2026.20260717_203001.bak.xlsx»
_BAK_RE = re.compile(r"^.+\.\d{8}_\d{6}\.bak\.xlsx$", re.IGNORECASE)

BACKUP_DIR_NAME = "backups"
DEFAULT_BACKUP_KEEP = 20


def backups_dir(summary_path: str | Path) -> Path:
    """Папка бэкапов рядом со сводной (не в корне приложения как россыпь файлов)."""
    return Path(summary_path).resolve().parent / BACKUP_DIR_NAME


def backup_filename(summary_path: str | Path, stamp: Optional[str] = None) -> str:
    p = Path(summary_path)
    stamp = stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{p.stem}.{stamp}.bak{p.suffix}"


def _is_backup_for(summary_path: Path, file: Path) -> bool:
    name = file.name
    if not name.lower().endswith(".bak.xlsx"):
        return False
    stem = summary_path.stem
    if name.startswith(stem + ".") and _BAK_RE.match(name):
        return True
    if stem in name and ".bak." in name.lower():
        return True
    return False


def list_backups(summary_path: str | Path, limit: int = 50) -> List[Path]:
    """
    Список бэкапов (новые → старые).
    Смотрит папку backups/ и для совместимости — старые файлы рядом со сводной.
    """
    p = Path(summary_path).resolve()
    found: List[Path] = []
    seen = set()

    dirs = [backups_dir(p)]
    if p.parent.exists():
        dirs.append(p.parent)

    for folder in dirs:
        if not folder.exists() or not folder.is_dir():
            continue
        for f in folder.iterdir():
            if not f.is_file() or not _is_backup_for(p, f):
                continue
            key = f.resolve()
            if key in seen:
                continue
            seen.add(key)
            found.append(f)

    found.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return found[:limit]


def rotate_backups(summary_path: str | Path, keep: int = DEFAULT_BACKUP_KEEP) -> List[Path]:
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


def make_backup(
    summary_path: str | Path,
    *,
    keep: int = DEFAULT_BACKUP_KEEP,
) -> Optional[Path]:
    """
    Копирует сводную в backups/{stem}.{stamp}.bak.xlsx и ротирует до keep файлов.
    Возвращает путь к новому бэкапу или None, если исходника нет.
    """
    src = Path(summary_path).resolve()
    if not src.exists():
        return None
    dest_dir = backups_dir(src)
    dest_dir.mkdir(parents=True, exist_ok=True)
    bak = dest_dir / backup_filename(src)
    shutil.copy2(src, bak)
    rotate_backups(src, keep=keep)
    return bak


def restore_backup(backup_path: str | Path, summary_path: str | Path) -> Path:
    """Копирует бэкап поверх сводной (текущий файл сначала уходит в backups/)."""
    bak = Path(backup_path)
    dest = Path(summary_path)
    if not bak.exists():
        raise FileNotFoundError(bak)
    if dest.exists():
        make_backup(dest, keep=DEFAULT_BACKUP_KEEP)
    shutil.copy2(bak, dest)
    return dest
