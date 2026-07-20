# analyzers/release_notes.py
"""Разбор RELEASE_NOTES.md и текст «Что нового» для версии."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

APP_DIR = Path(__file__).resolve().parent.parent
DEFAULT_NOTES = APP_DIR / "RELEASE_NOTES.md"

_SECTION_RE = re.compile(r"^##\s+(.+?)\s*$", re.MULTILINE)


def load_release_notes(path: Optional[Path] = None) -> str:
    p = Path(path) if path else DEFAULT_NOTES
    if not p.exists():
        return ""
    return p.read_text(encoding="utf-8")


def parse_release_sections(text: str) -> Dict[str, str]:
    """
    Секции вида «## 1.0.3» → текст пунктов.
    Ключ — номер версии без префикса v.
    """
    if not text:
        return {}
    matches = list(_SECTION_RE.finditer(text))
    sections: Dict[str, str] = {}
    for i, m in enumerate(matches):
        title = m.group(1).strip()
        ver = title
        if ver.lower().startswith("v"):
            ver = ver[1:].strip()
        # отсекаем не-версии (если появятся)
        if not re.match(r"^\d+(\.\d+)*", ver):
            continue
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = text[start:end].strip()
        sections[ver] = body
    return sections


def notes_for_version(version: str, path: Optional[Path] = None) -> str:
    ver = str(version or "").strip()
    if ver.lower().startswith("v"):
        ver = ver[1:].strip()
    sections = parse_release_sections(load_release_notes(path))
    return sections.get(ver, "")


def format_whats_new(
    version: str,
    *,
    path: Optional[Path] = None,
    previous_version: Optional[str] = None,
) -> str:
    """
    Текст окна «Что нового».
    Если previous_version задан — собирает секции (previous, current] по порядку файла.
    """
    ver = str(version or "").strip().lstrip("vV")
    prev = (previous_version or "").strip().lstrip("vV") or None
    sections = parse_release_sections(load_release_notes(path))
    if not sections:
        return f"Установлена версия {ver}.\n\nПодробности — в RELEASE_NOTES.md на GitHub."

    # порядок как в файле (новые сверху)
    order = list(sections.keys())

    if prev and prev in sections and ver in sections and prev != ver:
        # взять все секции от ver включительно до prev исключительно
        try:
            i_new = order.index(ver)
            i_old = order.index(prev)
        except ValueError:
            i_new = i_old = -1
        if i_new >= 0 and i_old >= 0 and i_new < i_old:
            picked = order[i_new:i_old]
        else:
            picked = [ver] if ver in sections else []
    else:
        picked = [ver] if ver in sections else []

    if not picked:
        # нет точной секции — показать самую верхнюю
        if order:
            picked = [order[0]]
        else:
            return f"Версия {ver}."

    parts: List[str] = [f"Что нового в версии {ver}", ""]
    for v in picked:
        if len(picked) > 1:
            parts.append(f"## {v}")
            parts.append("")
        body = sections.get(v, "").strip()
        if body:
            parts.append(body)
            parts.append("")
    return "\n".join(parts).strip() + "\n"
