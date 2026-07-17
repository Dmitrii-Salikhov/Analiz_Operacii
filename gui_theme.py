# gui_theme.py — минимальная тема для macOS Tk 8.5 (без кастомных цветов кнопок)
from __future__ import annotations

from tkinter import ttk


def apply_ttk_style(root):
    """На macOS оставляем aqua — кастомный clam с цветами даёт «пустые» кнопки."""
    style = ttk.Style(root)
    # Не форсируем clam: на системном Tk 8.5.9 цветные стили часто невидимы.
    try:
        # чуть плотнее вкладки, без смены темы
        style.configure("TNotebook.Tab", padding=(12, 6))
        style.configure("Treeview", rowheight=22)
        style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"))
        style.configure("TButton", padding=(8, 4))
        style.configure("Toolbar.TButton", padding=(10, 6))
        style.configure("KPI.TLabel", font=("Helvetica", 18, "bold"))
        style.configure("KPITitle.TLabel", font=("Helvetica", 9))
        style.configure("Header.TLabel", font=("Helvetica", 16, "bold"))
        style.configure("Status.TLabel", font=("Helvetica", 9))
    except Exception:
        pass
    return style
