# analyzers/file_lock.py
"""Проверка, что Excel-файл реально нельзя перезаписать (не путать с «хвостом» ~$)."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


class FileLockedError(RuntimeError):
    """Целевой файл нельзя перезаписать."""

    def __init__(
        self,
        path: str | Path,
        *,
        hint: str = "",
        reason: str = "",
        stale_lock: Optional[Path] = None,
    ):
        self.path = str(path)
        self.reason = reason or "locked"
        self.stale_lock = Path(stale_lock) if stale_lock else None
        if hint:
            self.hint = hint
        elif self.stale_lock is not None:
            self.hint = (
                f"Рядом лежит служебный файл Excel «{self.stale_lock.name}» "
                "(часто остаётся после сбоя). Его можно удалить — сам Excel при этом "
                "может быть уже закрыт."
            )
        else:
            self.hint = (
                "Файл сейчас нельзя открыть на запись (Excel, облако OneDrive/Диск, "
                "антивирус или права доступа). Закройте сводную, подождите синхронизацию "
                "и нажмите «Повторить»."
            )
        super().__init__(f"Файл занят / недоступен для записи:\n{self.path}\n\n{self.hint}")


@dataclass
class LockStatus:
    locked: bool
    reason: str = ""  # ok | missing | stale_lock_only | open_denied | no_write_perm
    lock_sibling: Optional[Path] = None
    detail: str = ""


def excel_lock_sibling(path: str | Path) -> Path:
    p = Path(path)
    return p.with_name(f"~${p.name}")


def remove_stale_excel_lock(path: str | Path) -> bool:
    """Удаляет ~$… если есть. True, если удалили."""
    lock = excel_lock_sibling(path)
    if not lock.exists():
        return False
    lock.unlink(missing_ok=True)
    return True


def probe_excel_lock(path: str | Path) -> LockStatus:
    """
    Реальная проверка: можно ли открыть файл на запись.
    Наличие ~$ само по себе НЕ считается блокировкой (это частый «хвост» после сбоя Excel).
    """
    p = Path(path)
    if not p.exists():
        return LockStatus(locked=False, reason="missing")

    lock = excel_lock_sibling(p)
    has_lock = lock.exists()

    if not os.access(p, os.W_OK):
        return LockStatus(
            locked=True,
            reason="no_write_perm",
            lock_sibling=lock if has_lock else None,
            detail="нет права на запись",
        )

    try:
        # r+b — попытка записи без усечения; на Windows падает, если Excel держит файл
        with open(p, "r+b"):
            pass
    except PermissionError as e:
        return LockStatus(
            locked=True,
            reason="open_denied",
            lock_sibling=lock if has_lock else None,
            detail=str(e),
        )
    except OSError as e:
        # WinError 32 и аналоги
        win = getattr(e, "winerror", None)
        if win == 32 or e.errno in (13, 11, 16):
            return LockStatus(
                locked=True,
                reason="open_denied",
                lock_sibling=lock if has_lock else None,
                detail=str(e),
            )
        return LockStatus(
            locked=True,
            reason="open_denied",
            lock_sibling=lock if has_lock else None,
            detail=str(e),
        )

    # Файл открывается — не занят. ~$ при этом может быть устаревшим.
    return LockStatus(
        locked=False,
        reason="stale_lock_only" if has_lock else "ok",
        lock_sibling=lock if has_lock else None,
        detail="есть ~$ но файл доступен для записи" if has_lock else "",
    )


def excel_file_locked(path: str | Path) -> bool:
    """True только если файл реально нельзя открыть на запись."""
    return probe_excel_lock(path).locked


def is_writable(path: str | Path) -> bool:
    p = Path(path)
    if not p.exists():
        return True
    return os.access(p, os.W_OK) and not excel_file_locked(path)


def ensure_excel_writable(path: str | Path) -> None:
    """Бросает FileLockedError, если сводную нельзя перезаписать."""
    p = Path(path)
    st = probe_excel_lock(p)
    if not st.locked:
        return
    raise FileLockedError(
        p,
        reason=st.reason,
        stale_lock=st.lock_sibling if st.reason != "open_denied" else st.lock_sibling,
        hint=(
            (
                f"ОС не даёт записать в файл ({st.detail or st.reason}). "
                "Частые причины: файл открыт в Excel; синхронизация OneDrive; "
                "антивирус временно держит файл."
            )
            if st.reason in ("open_denied", "no_write_perm")
            else ""
        ),
    )
