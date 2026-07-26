# ui_flet/screens/home.py
"""Главная оболочка: реестр экранов/действий из schemas/app_registry.yaml."""
from __future__ import annotations

from typing import Callable

import flet as ft

from ui_flet.app_context import AppContext
from ui_flet.schema_loader import actions_checklist, load_registry, screens_by_status


def _status_color(st: str):
    return {
        "done": ft.Colors.GREEN,
        "shell": ft.Colors.ORANGE,
        "backend": ft.Colors.BLUE_GREY,
    }.get(st, ft.Colors.GREY)


def _stub(page: ft.Page, screen_id: str) -> None:
    try:
        page.show_dialog(
            ft.SnackBar(
                content=ft.Text(
                    f"Экран «{screen_id}» ещё shell — доступен в Tk (app_desktop.py). "
                    "См. docs/INVENTORY_FLET.md"
                )
            )
        )
    except Exception:
        pass


def _hint_tk(page: ft.Page) -> None:
    try:
        page.show_dialog(ft.SnackBar(content=ft.Text("Классический UI: python3 app_desktop.py")))
    except Exception:
        pass


def build_home(page: ft.Page, ctx: AppContext, *, on_open_form14: Callable) -> ft.Control:
    reg = load_registry()
    by_st = screens_by_status(reg)
    actions = actions_checklist(reg)

    screen_tiles = []
    for st in ("done", "shell", "backend"):
        for s in by_st.get(st) or []:
            sid = str(s.get("id") or "")
            title = str(s.get("title") or sid)

            def _click(_e, i=sid):
                if i == "form14":
                    on_open_form14()
                else:
                    _stub(page, i)

            screen_tiles.append(
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.CIRCLE, color=_status_color(st), size=14),
                    title=ft.Text(title),
                    subtitle=ft.Text(f"{sid} · {st}"),
                    on_click=_click,
                )
            )

    done_n = sum(1 for a in actions if a.get("status") == "done")
    shell_n = sum(1 for a in actions if a.get("status") == "shell")
    action_rows = [
        ft.Text(f"Действия меню: done={done_n}, shell={shell_n}, всего={len(actions)}", size=12)
    ]
    for a in actions:
        st = str(a.get("status") or "shell")
        action_rows.append(
            ft.Text(f"• [{st}] {a.get('title')}", size=12, color=_status_color(st))
        )

    return ft.Column(
        [
            ft.Text(f"Сводная операций  v{ctx.version}", size=22, weight=ft.FontWeight.BOLD),
            ft.Text(
                "Flet + схемные параметры. Tk (app_desktop.py) — fallback, пока экраны shell.",
                size=13,
                color=ft.Colors.ON_SURFACE_VARIANT,
            ),
            ft.FilledButton("Открыть конструктор ФСН 14", on_click=lambda e: on_open_form14()),
            ft.OutlinedButton("Подсказка: классический Tk UI", on_click=lambda e: _hint_tk(page)),
            ft.Divider(),
            ft.Text("Экраны (schemas/app_registry.yaml)", weight=ft.FontWeight.BOLD),
            ft.Column(screen_tiles, scroll=ft.ScrollMode.AUTO, height=280),
            ft.Divider(),
            ft.Text("Чеклист действий", weight=ft.FontWeight.BOLD),
            ft.Column(action_rows, scroll=ft.ScrollMode.AUTO, height=220),
            ft.Text(
                "Инвентарь: docs/INVENTORY_FLET.md · схемы: schemas/",
                size=11,
                color=ft.Colors.ON_SURFACE_VARIANT,
            ),
        ],
        expand=True,
        scroll=ft.ScrollMode.AUTO,
        spacing=8,
    )
