# ui_flet/app.py
"""Полноценное Flet-приложение: тот же backend, что Tk, + конструктор ФСН 14."""
from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import flet as ft

from analyzers.file_lock import FileLockedError, remove_stale_excel_lock
from analyzers.summary_writer import MONTH_RU
from ui_flet.app_context import APP_ROOT, AppContext
from ui_flet.screens.form14 import Form14ConstructorView
from ui_flet.session import AppSession

NAV = [
    ("work", "Работа", ft.Icons.HOME),
    ("preview", "Превью", ft.Icons.TABLE_CHART),
    ("checks", "Проверки", ft.Icons.FACT_CHECK),
    ("emk", "ЭМК", ft.Icons.COMPARE_ARROWS),
    ("uncl", "Не класс.", ft.Icons.HELP_OUTLINE),
    ("disp", "Спорные", ft.Icons.WARNING_AMBER),
    ("form14", "ФСН 14", ft.Icons.EDIT_NOTE),
    ("log", "Журнал", ft.Icons.ARTICLE),
]

NAV_TIP = {
    "work": "Источники, настройки и сводка по загруженным данным",
    "preview": "Таблица по неделям: категории, итоги, форма 4001",
    "checks": "Длительные операции и отсутствие занесения на опер. стол",
    "emk": "Сверка план/экстренные с выгрузкой ЭМК",
    "uncl": "Операции без категории в рубрикаторе",
    "disp": "Операции с несколькими кандидатами категории",
    "form14": "Конструктор соответствия кодов строкам формы № 14",
    "log": "Журнал работы приложения (analysis.log)",
}


class AnalizApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.session = AppSession(log=self._on_log_line)
        self.log_lines: List[str] = []
        self.current_key = "work"
        self._preview_section = 0  # 0 категории / 1 итоги / 2 форма 4001
        self.nav_index = 0
        self.status = ft.Text(self.session.status, size=12)
        self.body = ft.Container(
            expand=True,
            padding=16,
            bgcolor=ft.Colors.SURFACE,
            alignment=ft.Alignment.TOP_LEFT,
            content=ft.Text("Загрузка…"),
        )
        self.file_picker = ft.FilePicker()
        try:
            page.services.append(self.file_picker)
        except Exception:
            page.overlay.append(self.file_picker)

        page.title = f"Сводная операций  v{self.session.version}"
        page.window.width = 1280
        page.window.height = 820
        page.window.min_width = 960
        page.window.min_height = 640
        for _name in ("app_icon.ico", "app_icon.png"):
            for _base in (
                APP_ROOT / "assets",
                APP_ROOT / "_internal" / "assets",
            ):
                _icon = _base / _name
                if _icon.is_file():
                    try:
                        page.window.icon = str(_icon)
                    except Exception:
                        pass
                    break
            else:
                continue
            break
        page.padding = 0
        page.bgcolor = ft.Colors.SURFACE

        self._build_chrome()
        self.sidebar = self._build_sidebar()
        page.add(
            ft.Column(
                [
                    self.top_bar,
                    self.warn_banner,
                    ft.Row(
                        [
                            self.sidebar,
                            ft.VerticalDivider(width=1),
                            self.body,
                        ],
                        expand=True,
                        spacing=0,
                        vertical_alignment=ft.CrossAxisAlignment.STRETCH,
                    ),
                    ft.Container(
                        content=self.status,
                        padding=ft.Padding(left=12, right=12, top=6, bottom=6),
                        bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST,
                    ),
                ],
                expand=True,
                spacing=0,
            )
        )
        self.show("work")

    def _on_log_line(self, line: str) -> None:
        self.log_lines.append(line)
        if len(self.log_lines) > 500:
            self.log_lines = self.log_lines[-500:]

    def _snack(self, msg: str) -> None:
        try:
            self.page.show_dialog(ft.SnackBar(content=ft.Text(msg), open=True))
        except Exception:
            self.status.value = msg
            try:
                self.page.update()
            except Exception:
                pass

    def _set_status(self, msg: Optional[str] = None) -> None:
        self.status.value = msg or self.session.status
        self._refresh_warn_banner()

    def _class_issue_counts(self) -> tuple[int, int]:
        n_uncl = len(self.session.unclassified_rows or [])
        n_disp = len(self.session.disputed_rows or [])
        return n_uncl, n_disp

    def _refresh_warn_banner(self) -> None:
        n_uncl, n_disp = self._class_issue_counts()
        if n_uncl <= 0 and n_disp <= 0:
            self.warn_banner.visible = False
            self.warn_banner.content = ft.Container()
            return

        parts: List[str] = []
        if n_uncl:
            parts.append(f"не классифицировано: {n_uncl}")
        if n_disp:
            parts.append(f"спорных: {n_disp}")
        msg = "Перед записью в Excel разберите: " + ", ".join(parts)

        actions: List[ft.Control] = [
            ft.Icon(ft.Icons.WARNING_AMBER, color=ft.Colors.ON_SECONDARY_CONTAINER, size=18),
            ft.Text(msg, size=12, weight=ft.FontWeight.BOLD, expand=True),
        ]
        if n_uncl:
            actions.append(
                ft.TextButton(
                    f"Не класс. ({n_uncl})",
                    tooltip="Открыть конструктор неклассифицированных",
                    on_click=lambda e: self.show("uncl"),
                )
            )
        if n_disp:
            actions.append(
                ft.TextButton(
                    f"Спорные ({n_disp})",
                    tooltip="Назначить категорию спорным операциям",
                    on_click=lambda e: self.show("disp"),
                )
            )

        self.warn_banner.content = ft.Row(actions, spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        self.warn_banner.visible = True

    def _build_chrome(self) -> None:
        self.dept_dd = ft.Dropdown(
            label="Отделение",
            width=260,
            dense=True,
            options=[ft.dropdown.Option(d) for d in self.session.departments()],
            value=self.session.department,
            on_select=self._on_dept,
            tooltip="Текущее отделение для фильтра журнала и сводной",
        )
        self.version_label = ft.Text(f"v{self.session.version}", weight=ft.FontWeight.BOLD)
        self.top_bar = ft.Container(
            content=ft.Row(
                [
                    self.version_label,
                    self.dept_dd,
                    ft.FilledButton(
                        "Опержурнал…",
                        tooltip="Загрузить один или несколько файлов опержурнала",
                        on_click=self._pick_surg,
                    ),
                    ft.OutlinedButton(
                        "Из папки…",
                        tooltip="Загрузить все Excel/CSV опержурналы из выбранной папки",
                        on_click=self._pick_surg_folder,
                    ),
                    ft.OutlinedButton(
                        "ЭМК…",
                        tooltip="Загрузить выгрузку ЭМК для сверки план/экстренные",
                        on_click=self._pick_emk,
                    ),
                    ft.FilledButton(
                        "Обновить",
                        tooltip="Пересчитать анализ и превью по загруженным данным",
                        on_click=lambda e: self._run_analysis(),
                    ),
                    ft.FilledButton(
                        "В Excel…",
                        tooltip="Записать недели и форму 4001 в файл сводной",
                        on_click=self._write_excel,
                    ),
                    ft.OutlinedButton(
                        "Отчёт…",
                        tooltip="Сохранить простой отчёт по текущему месяцу превью",
                        on_click=self._export_report,
                    ),
                    ft.OutlinedButton(
                        "Открыть Excel",
                        tooltip="Открыть файл сводной во внешнем Excel",
                        on_click=self._open_excel,
                    ),
                    ft.OutlinedButton(
                        "Обновления…",
                        tooltip="Проверить обновления на GitHub Releases",
                        on_click=self._check_updates,
                    ),
                ],
                spacing=8,
                scroll=ft.ScrollMode.AUTO,
            ),
            padding=ft.Padding(left=12, right=12, top=10, bottom=10),
            bgcolor=ft.Colors.SURFACE_CONTAINER,
        )
        self.warn_banner = ft.Container(
            visible=False,
            padding=ft.Padding(left=12, right=12, top=6, bottom=6),
            bgcolor=ft.Colors.SECONDARY_CONTAINER,
            content=ft.Container(),
        )

    def _kpi_cards(self) -> ft.Control:
        k = self.session.kpi
        items = [
            ("Операций", k["ops"]),
            ("Пациентов", k["patients"]),
            ("План", k["plan"]),
            ("Экстренных", k["emerg"]),
            ("Период", k["period"]),
            ("Файлы / ЭМК", k["files"]),
            ("Расхождений ЭМК", k["diff"]),
            ("Проверок", k.get("checks", "—")),
        ]

        def card(title: str, value: str) -> ft.Control:
            return ft.Container(
                content=ft.Column(
                    [
                        ft.Text(title, size=11, color=ft.Colors.ON_SURFACE_VARIANT),
                        ft.Text(str(value), size=16, weight=ft.FontWeight.BOLD),
                    ],
                    spacing=2,
                ),
                padding=ft.Padding(left=12, right=12, top=10, bottom=10),
                bgcolor=ft.Colors.SURFACE_CONTAINER_HIGH,
                border_radius=8,
            )

        return ft.Column(
            [
                ft.Text("Сводка", size=16, weight=ft.FontWeight.BOLD),
                ft.Row([card(t, v) for t, v in items], wrap=True, spacing=8, run_spacing=8),
            ],
            spacing=8,
        )

    def _build_sidebar(self) -> ft.Control:
        n_uncl, n_disp = self._class_issue_counts()

        def mk(i: int, key: str, lab: str, ic):
            selected = i == self.nav_index
            label = lab
            if key == "uncl" and n_uncl:
                label = f"{lab} ({n_uncl})"
            elif key == "disp" and n_disp:
                label = f"{lab} ({n_disp})"
            warn = (key == "uncl" and n_uncl > 0) or (key == "disp" and n_disp > 0)

            def go(_e=None, _i=i, _k=key):
                self.nav_index = _i
                self.show(_k)

            return ft.Container(
                content=ft.Row(
                    [
                        ft.Icon(
                            ic,
                            size=18,
                            color=ft.Colors.SECONDARY if warn and not selected else None,
                        ),
                        ft.Text(
                            label,
                            size=12,
                            weight=ft.FontWeight.BOLD if selected or warn else None,
                            color=ft.Colors.SECONDARY if warn and not selected else None,
                        ),
                    ],
                    spacing=8,
                ),
                padding=ft.Padding(left=10, right=10, top=8, bottom=8),
                bgcolor=ft.Colors.PRIMARY_CONTAINER if selected else (
                    ft.Colors.SECONDARY_CONTAINER if warn else None
                ),
                border_radius=8,
                ink=True,
                on_click=go,
                tooltip=NAV_TIP.get(key, lab),
            )

        items = [mk(i, key, lab, ic) for i, (key, lab, ic) in enumerate(NAV)]
        return ft.Container(
            width=130,
            padding=8,
            bgcolor=ft.Colors.SURFACE_CONTAINER_LOW,
            content=ft.Column(items, spacing=4, scroll=ft.ScrollMode.AUTO, expand=True),
        )

    def _refresh_sidebar(self) -> None:
        self.sidebar.content = self._build_sidebar().content
        self.sidebar.bgcolor = ft.Colors.SURFACE_CONTAINER_LOW

    def _on_dept(self, e) -> None:
        self.session.set_department(self.dept_dd.value or self.session.department)
        self._snack(f"Отделение: {self.session.department}")
        if not self.session.store.ops.empty:
            self._run_analysis()

    def show(self, key: str) -> None:
        builders = {
            "work": self._screen_work,
            "preview": self._screen_preview,
            "checks": self._screen_checks,
            "emk": self._screen_emk,
            "uncl": self._screen_uncl,
            "disp": self._screen_disp,
            "form14": self._screen_form14,
            "log": self._screen_log,
        }
        self.current_key = key
        for i, (k, _, _) in enumerate(NAV):
            if k == key:
                self.nav_index = i
                break
        try:
            self.body.content = builders.get(key, self._screen_work)()
        except Exception:
            import traceback

            tb = traceback.format_exc()
            self.session.log(tb, level="ERROR")
            self.body.content = ft.Column(
                [
                    ft.Text(f"Ошибка экрана «{key}»", color=ft.Colors.ERROR, size=18),
                    ft.Text(tb, selectable=True, size=11),
                ],
                scroll=ft.ScrollMode.AUTO,
                expand=True,
            )
        self._refresh_sidebar()
        self._set_status()
        self.page.update()

    # ——— actions ———
    async def _pick_surg(self, _e=None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="Опержурнал(ы)",
            allow_multiple=True,
            allowed_extensions=["xlsx", "xls", "csv"],
            file_type=ft.FilePickerFileType.CUSTOM,
            initial_directory=self.session.last_surg_dir,
        )
        if not files:
            return
        paths = [f.path for f in files if getattr(f, "path", None)]
        if not paths:
            self.session.log(
                "FilePicker не вернул path (файлы: "
                + ", ".join(getattr(f, "name", "?") for f in files)
                + ")",
                level="ERROR",
            )
            self._snack("Не удалось получить пути к файлам")
            return
        self.session.last_surg_dir = str(Path(paths[0]).parent)
        try:
            self.session.ingest_surg_paths(paths)
            self._snack(f"Загружено файлов: {len(paths)}")
            self.nav_index = 1
            self.show("preview")
        except Exception as ex:
            import traceback

            self.session.log(traceback.format_exc(), level="ERROR")
            self._snack(str(ex))

    async def _pick_surg_folder(self, _e=None) -> None:
        folder = await self.file_picker.get_directory_path(
            dialog_title="Папка с опержурналами",
            initial_directory=self.session.last_surg_dir,
        )
        if not folder:
            return
        base = Path(folder)
        self.session.last_surg_dir = folder
        paths = sorted(
            [*(base.glob("*.xlsx")), *(base.glob("*.xls")), *(base.glob("*.csv"))],
            key=lambda p: p.name.lower(),
        )
        paths = [p for p in paths if not p.name.startswith("~$") and ".bak." not in p.name.lower()]
        if not paths:
            self._snack("В папке нет Excel/CSV")
            return
        try:
            self.session.ingest_surg_paths([str(p) for p in paths])
            self._snack(f"Из папки: {len(paths)} файлов")
            self.nav_index = 1
            self.show("preview")
        except Exception as ex:
            self._snack(str(ex))

    async def _pick_emk(self, _e=None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="ЭМК",
            allow_multiple=False,
            allowed_extensions=["xlsx", "xls", "csv"],
            initial_directory=self.session.last_emk_dir,
        )
        if not files or not getattr(files[0], "path", None):
            return
        try:
            self.session.load_emk_path(files[0].path)
            self._snack("ЭМК загружен")
            self.show("emk")
        except Exception as ex:
            self._snack(str(ex))

    def _run_analysis(self) -> None:
        try:
            self.session.run_analysis()
            self._snack(self.session.status)
            self.show(self.current_key)
        except Exception as ex:
            self._snack(str(ex))

    async def _write_excel(self, _e=None) -> None:
        if self.session.store.ops.empty:
            self._snack("Сначала загрузите опержурнал")
            return
        path = self.session.summary_path
        if not path or not Path(path).exists():
            files = await self.file_picker.pick_files(
                dialog_title="Выберите файл сводной",
                allow_multiple=False,
                allowed_extensions=["xlsx"],
                initial_directory=str(Path(path).parent if path else self.session.app_dir),
            )
            if not files or not files[0].path:
                return
            self.session.summary_path = files[0].path
            self.session.persist()

        await self._write_excel_attempt()

    async def _write_excel_attempt(self) -> None:
        write_form = self.session.write_form and self.session.form4001_enabled()
        path = self.session.summary_path
        try:
            report = self.session.write_excel(
                write_weeks=self.session.write_weeks,
                write_form=write_form,
            )
            msg = f"Записано ячеек: {report.get('cells_written', 0)}"
            if report.get("backup"):
                msg += f"\nБэкап: {Path(report['backup']).name}"
            if report.get("verify_msg"):
                msg += f"\n{report['verify_msg']}"
            self._snack(msg)
            self._set_status()
        except FileLockedError as ex:
            await self._show_file_locked_dialog(ex)
        except Exception as ex:
            self._snack(str(ex))

    async def _show_file_locked_dialog(self, err: FileLockedError) -> None:
        name = Path(err.path).name
        lock_name = err.stale_lock.name if err.stale_lock else ""

        async def retry(_e=None):
            dlg.open = False
            self.page.update()
            await self._write_excel_attempt()

        def close(_e=None):
            dlg.open = False
            self.page.update()

        def reveal(_e=None):
            try:
                import os
                import subprocess

                if os.name == "nt":
                    subprocess.Popen(["explorer", "/select,", err.path])
                else:
                    subprocess.Popen(["open", str(Path(err.path).parent)])
            except Exception:
                pass

        async def clear_lock_and_retry(_e=None):
            try:
                remove_stale_excel_lock(err.path)
            except OSError as ex:
                self._snack(f"Не удалось удалить ~$: {ex}")
            dlg.open = False
            self.page.update()
            await self._write_excel_attempt()

        actions = [
            ft.TextButton("Показать в папке", on_click=reveal),
            ft.TextButton("Отмена", on_click=close),
        ]
        if err.stale_lock is not None:
            actions.insert(
                0,
                ft.OutlinedButton(
                    f"Удалить {lock_name}",
                    tooltip="Убрать служебный «хвост» Excel после сбоя",
                    on_click=clear_lock_and_retry,
                ),
            )
        actions.append(ft.FilledButton("Повторить", on_click=retry))

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Не удалось записать сводную"),
            content=ft.Text(
                f"Файл: «{name}»\n\n{err.hint}\n\n{err.path}",
                selectable=True,
            ),
            actions=actions,
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(dlg)

    async def _export_report(self, _e=None) -> None:
        if self.session.store.ops.empty:
            self._snack("Нет данных")
            return
        month = self.session.month_label_to_num.get(self.session.preview_month)
        default = f"Отчёт_{MONTH_RU.get(month or 0, 'период')}_{self.session.year}.xlsx"
        path = await self.file_picker.save_file(
            dialog_title="Экспорт отчёта",
            file_name=default,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            initial_directory=str(self.session.app_dir),
        )
        if not path:
            path = str(self.session.app_dir / default)
        try:
            self.session.export_simple_report(path)
            self._snack(f"Отчёт: {path}")
        except Exception as ex:
            self._snack(str(ex))

    def _open_excel(self, _e=None) -> None:
        p = self.session.summary_path
        if not p or not Path(p).exists():
            self._snack("Файл сводной не найден")
            return
        self.session.open_path(p)

    async def _check_updates(self, _e=None) -> None:
        import asyncio

        busy = ft.AlertDialog(
            modal=True,
            title=ft.Text("Обновления"),
            content=ft.Column(
                [
                    ft.ProgressBar(width=360),
                    ft.Text("Проверка GitHub Releases…", size=12),
                ],
                tight=True,
                width=380,
            ),
            actions=[],
        )
        self.page.show_dialog(busy)
        try:
            info = await asyncio.to_thread(self.session.check_for_app_update)
        except Exception as ex:
            busy.open = False
            self.page.update()
            self._snack(str(ex))
            return
        busy.open = False
        self.page.update()

        if info is None:
            self._snack(f"У вас актуальная версия: {self.session.version}")
            return

        notes = self.session.format_app_update_notes(info)
        include_cfg = ft.Checkbox(
            label="Также заменить config.yaml (обычно не нужно)",
            value=False,
        )
        notes_box = ft.TextField(
            value=notes,
            multiline=True,
            min_lines=10,
            max_lines=14,
            read_only=True,
            text_size=12,
            width=520,
        )

        offer = ft.AlertDialog(modal=True, title=ft.Text("Доступно обновление"), actions=[])

        def close_offer(_e=None):
            offer.open = False
            self.page.update()

        def open_github(_e=None):
            if info.html_url:
                self.session.open_path(info.html_url)

        async def do_install(_e=None):
            offer.open = False
            self.page.update()
            await self._install_update(info, include_config=bool(include_cfg.value))

        offer.content = ft.Column(
            [
                ft.Text(
                    f"Найдена версия {info.remote_version} (сейчас {info.local_version})",
                    weight=ft.FontWeight.BOLD,
                ),
                notes_box,
                include_cfg,
            ],
            tight=True,
            width=540,
            scroll=ft.ScrollMode.AUTO,
        )
        offer.actions = [
            ft.TextButton("Позже", on_click=close_offer),
            ft.OutlinedButton("GitHub", on_click=open_github),
            ft.FilledButton("Установить", on_click=do_install),
        ]
        offer.actions_alignment = ft.MainAxisAlignment.END
        self.page.show_dialog(offer)

    async def _install_update(self, info, *, include_config: bool) -> None:
        import asyncio

        from analyzers.release_notes import format_whats_new

        status = ft.Text("Подготовка…", size=12)
        bar = ft.ProgressBar(width=420, value=0)
        prog = ft.AlertDialog(
            modal=True,
            title=ft.Text(f"Установка v{info.remote_version}"),
            content=ft.Column([bar, status], tight=True, width=440),
            actions=[],
        )
        self.page.show_dialog(prog)

        def on_progress(msg: str, frac) -> None:
            status.value = msg or ""
            if frac is None:
                bar.value = None  # indeterminate
            else:
                try:
                    bar.value = max(0.0, min(1.0, float(frac)))
                except (TypeError, ValueError):
                    bar.value = None
            try:
                self.page.update()
            except Exception:
                pass

        try:
            report = await asyncio.to_thread(
                self.session.apply_app_update,
                info,
                include_config=include_config,
                on_progress=on_progress,
            )
        except PermissionError as ex:
            prog.open = False
            self.page.update()
            self._snack(
                "Не удалось заменить файлы — закройте программу и установите обновление снова. "
                f"{ex}"
            )
            return
        except Exception as ex:
            prog.open = False
            self.page.update()
            self._snack(str(ex))
            return

        new_ver = str(report.get("new_version") or self.session.version)
        self.session.version = new_ver
        self.version_label.value = f"v{new_ver}"
        self.page.title = f"Сводная операций  v{new_ver}"

        whats = format_whats_new(
            new_ver,
            path=self.session.app_dir / "RELEASE_NOTES.md",
            previous_version=info.local_version,
        )
        status.value = f"Готово: v{new_ver} ({report.get('count', 0)} файлов)"
        bar.value = 1
        self.page.update()

        done = ft.AlertDialog(modal=True, title=ft.Text(f"Установлена версия {new_ver}"), actions=[])

        def close_done(_e=None):
            done.open = False
            prog.open = False
            self.page.update()

        def restart(_e=None):
            try:
                self.session.spawn_restart()
            except Exception as ex:
                self._snack(f"Не удалось перезапустить: {ex}")
                return
            try:
                self.page.window.destroy()
            except Exception:
                pass
            raise SystemExit(0)

        done.content = ft.Column(
            [
                ft.Text(whats or f"Версия {new_ver} установлена.", selectable=True, size=12),
                ft.Text("Перезапустите приложение, чтобы применить изменения.", size=12),
            ],
            tight=True,
            width=520,
            height=280,
            scroll=ft.ScrollMode.AUTO,
        )
        done.actions = [
            ft.TextButton("Позже", on_click=close_done),
            ft.FilledButton("Перезапустить", on_click=restart),
        ]
        done.actions_alignment = ft.MainAxisAlignment.END
        prog.open = False
        self.page.show_dialog(done)

    # ——— screens ———
    def _screen_work(self) -> ft.Control:
        s = self.session
        year_tf = ft.TextField(label="Год", value=str(s.year), width=90, dense=True)
        start_tf = ft.TextField(label="Дата с", value=s.start_date, width=120, dense=True)
        end_tf = ft.TextField(label="Дата по", value=s.end_date, width=120, dense=True)
        long_tf = ft.TextField(
            label="Длительная опер., ч",
            value=f"{s.long_op_hours():g}",
            width=140,
            dense=True,
            tooltip="Операции дольше этого порога попадают в «Проверки»",
        )
        sum_tf = ft.TextField(label="Сводная", value=s.summary_path, expand=True, dense=True)
        filt = ft.Switch(label="Фильтр дат", value=s.filter_enabled)
        hide = ft.Switch(label="Скрыть нули", value=s.hide_zeros)
        plan = ft.Dropdown(
            label="План/экстр",
            width=160,
            dense=True,
            options=[
                ft.dropdown.Option("template", "По шаблону"),
                ft.dropdown.Option("emk", "Сверка ЭМК"),
            ],
            value=s.plan_mode,
        )
        weeks_sw = ft.Switch(label="Писать недели", value=s.write_weeks)
        form_sw = ft.Switch(
            label="Писать форму 4001",
            value=s.write_form and s.form4001_enabled(),
            disabled=not s.form4001_enabled(),
        )
        sources = ft.Text(s.sources_text(), selectable=True)
        emk = ft.Text(s.emk_status_text())

        def apply(_e=None):
            try:
                s.year = int(year_tf.value or s.year)
            except ValueError:
                pass
            s.start_date = start_tf.value or s.start_date
            s.end_date = end_tf.value or s.end_date
            s.summary_path = sum_tf.value or s.summary_path
            s.filter_enabled = bool(filt.value)
            s.hide_zeros = bool(hide.value)
            s.plan_mode = plan.value or "template"
            s.write_weeks = bool(weeks_sw.value)
            s.write_form = bool(form_sw.value) if s.form4001_enabled() else False
            try:
                s.set_long_op_hours(float(str(long_tf.value or "4").replace(",", ".")))
            except ValueError:
                pass
            s.persist()
            if not s.store.ops.empty:
                s.run_analysis()
            self._snack("Настройки применены")
            self._set_status()

        async def choose_summary(_e=None):
            files = await self.file_picker.pick_files(
                dialog_title="Сводная",
                allow_multiple=False,
                allowed_extensions=["xlsx"],
            )
            if files and files[0].path:
                sum_tf.value = files[0].path
                self.page.update()

        return ft.Column(
            [
                ft.Text("Источники и настройки", size=20, weight=ft.FontWeight.BOLD),
                self._kpi_cards(),
                ft.Divider(),
                ft.Text(
                    "Загрузите опержурнал → Обновить → Превью → Записать в Excel / Отчёт. "
                    "Конструктор ФСН 14 — в разделе слева.",
                    size=13,
                    color=ft.Colors.ON_SURFACE_VARIANT,
                ),
                ft.Row([year_tf, start_tf, end_tf, long_tf, filt, hide, plan], wrap=True),
                ft.Row(
                    [
                        sum_tf,
                        ft.OutlinedButton(
                            "Обзор…",
                            tooltip="Выбрать файл сводной Excel",
                            on_click=choose_summary,
                        ),
                    ],
                    expand=True,
                ),
                ft.Row(
                    [
                        weeks_sw,
                        form_sw,
                        ft.FilledButton(
                            "Применить",
                            tooltip="Сохранить настройки и пересчитать анализ",
                            on_click=apply,
                        ),
                    ]
                ),
                ft.Divider(),
                ft.Text("Опержурналы", weight=ft.FontWeight.BOLD),
                ft.Container(content=sources, padding=8, bgcolor=ft.Colors.SURFACE_CONTAINER_LOW),
                emk,
                ft.Row(
                    [
                        ft.OutlinedButton(
                            "Очистить накопитель",
                            tooltip="Удалить все загруженные операции из памяти",
                            on_click=self._clear,
                        ),
                        ft.OutlinedButton(
                            "План/экстр по ЭМК → config",
                            tooltip="Записать классификацию план/экстренные из ЭМК в config.yaml",
                            on_click=self._classify_kinds,
                        ),
                    ]
                ),
            ],
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=10,
        )

    def _clear(self, _e=None) -> None:
        self.session.clear_store()
        self._snack("Очищено")
        self.show("work")

    def _classify_kinds(self, _e=None) -> None:
        try:
            kind = self.session.classify_kinds_from_emk()
            self._snack(
                f"Экстр.: {len(kind.get('emergency') or [])}, план: {len(kind.get('plan') or [])}"
            )
        except Exception as ex:
            self._snack(str(ex))

    def _screen_preview(self) -> ft.Control:
        s = self.session
        labels = list(s.month_label_to_num.keys()) or list(MONTH_RU.values())
        month_dd = ft.Dropdown(
            label="Месяц",
            width=160,
            dense=True,
            options=[ft.dropdown.Option(x) for x in labels],
            value=s.preview_month if s.preview_month in labels else (labels[-1] if labels else None),
            on_select=lambda e: self._change_month(month_dd.value),
            tooltip="Месяц для превью сводной",
        )
        pb = s.preview
        form_tab = "Форма 4001" if pb.form_kind == "4001" else "Форма № 14"
        sections = ["Категории", "Итоги", form_tab]
        seg = self._preview_section if 0 <= self._preview_section < 3 else 0

        def switch(i: int):
            def _h(_e=None):
                self._preview_section = i
                self.show("preview")

            return _h

        if seg == 0:
            content = self._table(
                ["Категория"] + list(pb.week_headers) + ["ИТОГ"],
                pb.cat_rows,
            )
        elif seg == 1:
            content = self._table(
                ["Показатель"] + list(pb.week_headers) + ["ИТОГ"],
                pb.tot_rows,
            )
        else:
            content = self._form_table(pb.form_rows, kind=pb.form_kind)

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Превью", size=20, weight=ft.FontWeight.BOLD),
                        month_dd,
                        ft.Text(pb.info or "Нет данных — загрузите журнал", size=12),
                    ],
                    wrap=True,
                ),
                ft.Row(
                    [
                        (
                            ft.FilledButton(
                                lab,
                                tooltip=f"Показать: {lab}",
                                on_click=switch(i),
                            )
                            if i == seg
                            else ft.OutlinedButton(
                                lab,
                                tooltip=f"Показать: {lab}",
                                on_click=switch(i),
                            )
                        )
                        for i, lab in enumerate(sections)
                    ],
                    spacing=8,
                ),
                ft.Container(content=content, expand=True),
            ],
            expand=True,
            spacing=10,
        )

    def _change_month(self, label: Optional[str]) -> None:
        if not label:
            return
        self.session.preview_month = label
        self.session.build_preview()
        self.show("preview")

    def _data_table(
        self,
        headers: List[str],
        rows: List[List],
        *,
        numeric_from: int = 1,
        highlight_total: bool = True,
    ) -> ft.Control:
        if not rows:
            return ft.Text("Нет строк")
        total_idx = None
        if highlight_total:
            for i, h in enumerate(headers):
                if str(h).strip().upper() in ("ИТОГ", "ВСЕГО"):
                    total_idx = i
                    break
            if total_idx is None and len(headers) > 1:
                # последняя числовая колонка — обычно итог
                total_idx = len(headers) - 1

        columns = []
        for i, h in enumerate(headers):
            is_tot = total_idx is not None and i == total_idx
            columns.append(
                ft.DataColumn(
                    label=ft.Text(
                        str(h),
                        weight=ft.FontWeight.BOLD,
                        size=12,
                    ),
                    numeric=(i >= numeric_from),
                )
            )
        data_rows = []
        for r in rows[:500]:
            cells = []
            for i, _h in enumerate(headers):
                val = r[i] if i < len(r) else ""
                is_tot = total_idx is not None and i == total_idx
                cells.append(
                    ft.DataCell(
                        ft.Text(
                            str(val),
                            size=12,
                            weight=ft.FontWeight.BOLD if (i == 0 or is_tot) else None,
                        )
                    )
                )
            data_rows.append(ft.DataRow(cells=cells))
        table = ft.DataTable(
            columns=columns,
            rows=data_rows,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=8,
            heading_row_color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            data_row_min_height=36,
            heading_row_height=40,
            column_spacing=16,
            horizontal_lines=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT),
            show_checkbox_column=False,
        )
        return ft.Column(
            [ft.Row([table], scroll=ft.ScrollMode.AUTO)],
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

    def _table(self, headers: List[str], rows: List[List]) -> ft.Control:
        return self._data_table(headers, rows, numeric_from=1)

    def _form_table(self, rows: List[dict], *, kind: str = "4001") -> ft.Control:
        if not rows:
            if kind == "14":
                return ft.Text("Нет данных формы № 14 за выбранный месяц")
            return ft.Text("Нет данных формы 4001 (для не-ЛОР используйте превью формы № 14)")
        headers = [
            "Наименование",
            "Стр.",
            "Всего",
            "0–14 лет",
            "До 1 года",
            "15–17 лет",
            "Морфология",
            "Старше трудоспособного",
        ]
        data = [
            [
                r.get("name", ""),
                r.get("line", ""),
                r.get("total", ""),
                r.get("age_0_14", ""),
                r.get("age_under_1", ""),
                r.get("age_15_17", ""),
                r.get("histology", ""),
                r.get("senior", ""),
            ]
            for r in rows
        ]
        return self._data_table(headers, data, numeric_from=2)

    def _screen_checks(self) -> ft.Control:
        s = self.session
        long_rows = s.long_op_rows
        miss_rows = s.missing_table_rows
        thr = s.long_op_hours()

        def table_block(title: str, rows: list) -> ft.Control:
            headers = ["КВС", "Пациент", "Хирург", "Услуга", "Длит., ч", "Причина"]
            data = [
                [
                    r.get("КВС", ""),
                    r.get("Пациент", ""),
                    r.get("Хирург", ""),
                    (str(r.get("Услуга") or "")[:70]),
                    r.get("Длительность", "") if title.startswith("Длительн") else "",
                    r.get("Причина", ""),
                ]
                for r in rows[:500]
            ]
            body = (
                self._data_table(headers, data)
                if data
                else ft.Text("Нет замечаний", color=ft.Colors.ON_SURFACE_VARIANT)
            )
            return ft.Column(
                [
                    ft.Text(f"{title}: {len(rows)}", size=16, weight=ft.FontWeight.BOLD),
                    body,
                ],
                spacing=6,
            )

        async def exp(_e):
            path = await self.file_picker.save_file(
                dialog_title="Экспорт проверок",
                file_name="проверки_операций.xlsx",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx"],
            )
            if not path:
                path = str(self.session.app_dir / "проверки_операций.xlsx")
            n = self.session.export_quality_checks(path)
            self._snack(f"Экспорт проверок: {n} → {path}")

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Проверки журнала", size=20, weight=ft.FontWeight.BOLD),
                        ft.Container(expand=True),
                        ft.Text(f"Порог длительности: > {thr:g} ч", size=12),
                        ft.OutlinedButton(
                            "Экспорт…",
                            tooltip="Выгрузить длительные и без стола в Excel",
                            on_click=exp,
                        ),
                    ]
                ),
                ft.Text(
                    "№ истории, ФИО пациента, хирург, операция. "
                    "Порог меняется на вкладке «Работа».",
                    size=12,
                    color=ft.Colors.ON_SURFACE_VARIANT,
                ),
                table_block(f"Длительные операции (> {thr:g} ч)", long_rows),
                ft.Divider(),
                table_block("Не занесены на опер. стол", miss_rows),
            ],
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            spacing=10,
        )

    def _screen_emk(self) -> ft.Control:
        rows = self.session.emk_mismatch_rows
        lines = [f"Расхождений: {len(rows)}", self.session.emk_status_text(), ""]
        lines.append("Дата | КВС | Категория | Шаблон | ЭМК")
        for r in rows[:400]:
            lines.append(
                f"{r.get('Дата')} | {r.get('КВС')} | {r.get('Категория')} | "
                f"{r.get('Шаблон')} | {r.get('ЭМК')}"
            )

        async def exp(_e):
            path = await self.file_picker.save_file(
                file_name="расхождения_эмк.xlsx",
                allowed_extensions=["xlsx", "csv"],
            )
            if not path:
                return
            n = self.session.export_emk_mismatches(path)
            self._snack(f"Экспорт: {n} строк")

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Расхождения ЭМК", size=20, weight=ft.FontWeight.BOLD),
                        ft.OutlinedButton(
                            "Обновить сверку",
                            tooltip="Пересчитать расхождения план/экстренные с ЭМК",
                            on_click=lambda e: self._refresh_emk(),
                        ),
                        ft.OutlinedButton(
                            "Экспорт…",
                            tooltip="Сохранить таблицу расхождений в Excel/CSV",
                            on_click=exp,
                        ),
                    ]
                ),
                ft.Text("\n".join(lines), selectable=True, size=12),
            ],
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

    def _refresh_emk(self) -> None:
        self.session.refresh_emk_compare()
        self.show("emk")

    def _screen_uncl(self) -> ft.Control:
        async def exp(_e):
            path = await self.file_picker.save_file(
                file_name="неклассифицировано.xlsx",
                allowed_extensions=["xlsx", "csv"],
            )
            if not path:
                return
            n = self.session.export_unclassified(path)
            self._snack(f"Экспорт: {n}")

        async def problems(_e):
            path = await self.file_picker.save_file(
                file_name="проблемные_коды.xlsx",
                allowed_extensions=["xlsx", "csv"],
            )
            if not path:
                return
            n = self.session.export_problem_codes(path)
            self._snack(f"Проблемных кодов: {n}")

        def open_constructor(row: dict):
            # Делаем “конструктор” максимально близким по смыслу к Tk:
            # пользователь добавляет новую категорию в config.yaml, затем накопитель переклассифицируется.
            from analyzers.category_registry import (
                CategorySpec,
                FORM_LINES,
                suggest_keywords_from_name,
            )

            cats = list((self.session.summary_cfg.get("category_rows") or {}).keys())
            anchor_options = cats[:]
            anchor_default = cats[0] if cats else ""

            name_default = str(row.get("КСГ_название") or row.get("Услуга") or "").strip()[:120]
            codes_default = str(row.get("Код") or "").strip()
            if not codes_default:
                # иногда код в журнале отсутствует, тогда пользователь заполняет вручную
                codes_default = ""

            kw_default = ", ".join(suggest_keywords_from_name(name_default)) if name_default else ""

            name_field = ft.TextField(
                label="Название категории (как в отчёте)",
                value=name_default,
                width=540,
            )
            codes_field = ft.TextField(
                label="Код(ы) (через запятую)",
                value=codes_default,
                width=540,
            )
            kw_field = ft.TextField(
                label="Ключевые слова (через запятую)",
                value=kw_default,
                width=540,
            )
            kind_dd = ft.Dropdown(
                label="Тип",
                width=220,
                options=[
                    ft.dropdown.Option("plan", "Плановая"),
                    ft.dropdown.Option("emergency", "Экстренная"),
                ],
                value="plan",
            )
            line_dd = ft.Dropdown(
                label="Строка формы 4001",
                width=220,
                options=[ft.dropdown.Option(v, v) for v in FORM_LINES],
                value="6",
            )
            hist_cb = ft.Checkbox(label="Гистология", value=False)
            endo_cb = ft.Checkbox(label="Эндоскопия", value=False)
            anchor_dd = ft.Dropdown(
                label="Вставить после категории (якорь)",
                width=340,
                options=[ft.dropdown.Option(c, c) for c in anchor_options] if anchor_options else [],
                value=anchor_default if anchor_options else "",
            )

            dlg = ft.AlertDialog(modal=True, title=ft.Text("Конструктор: добавить категорию"), actions=[])

            def close(_e=None):
                dlg.open = False
                self.page.update()

            async def apply(_e=None):
                try:
                    def _split_csv(s: str) -> list[str]:
                        return [x.strip() for x in (s or "").replace(";", ",").split(",") if x.strip()]

                    spec = CategorySpec(
                        name=(name_field.value or "").strip(),
                        codes=_split_csv(codes_field.value or ""),
                        name_keywords=_split_csv(kw_field.value or ""),
                        kind=str(kind_dd.value or "plan"),
                        form_line=str(line_dd.value or "6"),
                        histology=bool(hist_cb.value),
                        endoscopic=bool(endo_cb.value),
                        anchor_category=str(anchor_dd.value or ""),
                    )
                    res = self.session.add_category_and_reclassify(spec)
                    self._snack(f"Категория добавлена: {res['added']}")
                    if res.get("warnings"):
                        self._snack("Есть предупреждения в лог/консоль")
                    close()
                    self.show("uncl")
                except Exception as ex:
                    self._snack(str(ex))

            actions = [
                ft.TextButton("Отмена", on_click=close),
                ft.FilledButton("Добавить и переклассифицировать", on_click=apply),
            ]
            dlg.title = ft.Text("Конструктор: добавить категорию")
            dlg.content = ft.Column(
                [
                    ft.Text(f"Операция: Дата {row.get('Дата')} | КВС {row.get('КВС')} | Код {row.get('Код')}", size=12),
                    name_field,
                    codes_field,
                    kw_field,
                    ft.Row([kind_dd, line_dd]),
                    ft.Row([hist_cb, endo_cb]),
                    anchor_dd,
                    ft.Text(
                        "После добавления категории накопитель переклассифицируется по ключевым словам.",
                        size=12,
                        color=ft.Colors.GREY_700,
                    ),
                ],
                width=600,
            )
            dlg.actions = actions
            dlg.actions_alignment = ft.MainAxisAlignment.END
            self.page.show_dialog(dlg)

        rows = self.session.unclassified_rows
        list_view = ft.ListView(
            expand=True,
            spacing=2,
            auto_scroll=True,
            controls=[
                ft.ListTile(
                    title=ft.Text(f"{r.get('Дата')} | КВС {r.get('КВС')}"),
                    subtitle=ft.Text(f"Код {r.get('Код')} | {r.get('КСГ_название')} | {r.get('Услуга')}"),
                    on_click=lambda e, r=r: open_constructor(r),
                )
                for r in rows[:500]
            ]
            or [ft.Text("(пусто)")],
        )

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Не классифицировано", size=20, weight=ft.FontWeight.BOLD),
                        ft.Text(f"({len(rows)})", size=12, color=ft.Colors.GREY_700),
                        ft.OutlinedButton(
                            "Экспорт…",
                            tooltip="Сохранить неклассифицированные операции в файл",
                            on_click=exp,
                        ),
                        ft.OutlinedButton(
                            "Проблемные коды…",
                            tooltip="Выгрузить коды без устойчивой категории",
                            on_click=problems,
                        ),
                    ]
                ),
                list_view,
            ],
            expand=True,
        )

    def _screen_disp(self) -> ft.Control:
        rows = self.session.disputed_rows

        def open_constructor(row: dict):
            from analyzers.surgery import lookup_category_meta

            all_cats = list((self.session.summary_cfg.get("category_rows") or {}).keys())
            ordered = []
            for c in (str(row.get("Кандидаты") or "").split("|") if row.get("Кандидаты") else []):
                c = str(c).strip()
                if c and c not in ordered:
                    ordered.append(c)
            for c in all_cats:
                if c and c not in ordered:
                    ordered.append(c)
            if row.get("Категория") and row.get("Категория") not in ordered:
                ordered.insert(0, str(row.get("Категория")))

            initial = str(row.get("Категория") or (ordered[0] if ordered else ""))

            cat_dd = ft.Dropdown(
                label="Категория для выбранной операции",
                width=520,
                options=[ft.dropdown.Option(c, c) for c in ordered] if ordered else [],
                value=initial if initial else None,
            )

            dlg = ft.AlertDialog(modal=True, title=ft.Text("Назначить категорию"), actions=[])

            def close(_e=None):
                dlg.open = False
                self.page.update()

            async def apply(_e=None):
                try:
                    cat = str(cat_dd.value or "").strip()
                    if not cat:
                        self._snack("Выберите категорию")
                        return
                    self.session.assign_disputed_category(int(row.get("StoreIndex")), cat)
                    self._snack(f"Спорные: назначено «{cat}»")
                    close()
                    self.show("disp")
                except Exception as ex:
                    self._snack(str(ex))

            dlg.content = ft.Column(
                [
                    ft.Text(
                        f"Операция: Дата {row.get('Дата')} | КВС {row.get('КВС')} | Код {row.get('Код')}",
                        size=12,
                    ),
                    cat_dd,
                    ft.Text(
                        "Назначение сохранится как ручное и не сбросится при правке ключей.",
                        size=12,
                        color=ft.Colors.GREY_700,
                    ),
                ],
                width=560,
            )
            dlg.actions = [
                ft.TextButton("Отмена", on_click=close),
                ft.FilledButton("Назначить", on_click=apply),
            ]
            dlg.actions_alignment = ft.MainAxisAlignment.END
            self.page.show_dialog(dlg)

        list_view = ft.ListView(
            expand=True,
            spacing=2,
            auto_scroll=True,
            controls=[
                ft.ListTile(
                    title=ft.Text(f"{r.get('Дата')} | КВС {r.get('КВС')}"),
                    subtitle=ft.Text(f"Код {r.get('Код')} | {r.get('Категория')} | кандидаты: {r.get('Кандидаты')}"),
                    on_click=lambda e, r=r: open_constructor(r),
                )
                for r in rows[:500]
            ]
            or [ft.Text("(пусто)")],
        )

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Спорные по ключам", size=20, weight=ft.FontWeight.BOLD),
                        ft.Text(f"({len(rows)})", size=12, color=ft.Colors.GREY_700),
                    ]
                ),
                list_view,
            ],
            expand=True,
        )

    def _screen_form14(self) -> ft.Control:
        ctx = AppContext(self.session.app_dir)
        ctx.config = self.session.config
        ctx.settings["department"] = self.session.department
        view = Form14ConstructorView(
            self.page,
            ctx,
            on_back=lambda: self.show("work"),
            default_dept_key=self.session.summary_key,
        )
        return view.build()

    def _screen_log(self) -> ft.Control:
        lines = self.session.log_lines_list() or list(self.log_lines)
        n = len(lines)
        status = ft.Text(f"Строк: {n} / 500 (старые удаляются)", size=12)

        list_view = ft.ListView(
            expand=True,
            spacing=2,
            auto_scroll=True,
            controls=[
                ft.Text(ln, selectable=True, size=11, font_family="Menlo")
                for ln in (lines or ["(пусто)"])
            ],
        )

        def refresh(_e=None):
            cur = self.session.log_lines_list() or list(self.log_lines)
            list_view.controls = [
                ft.Text(ln, selectable=True, size=11, font_family="Menlo")
                for ln in (cur or ["(пусто)"])
            ]
            status.value = f"Строк: {len(cur)} / 500 (старые удаляются)"
            list_view.auto_scroll = True
            self.page.update()
            try:
                list_view.scroll_to(offset=-1, duration=200)
            except Exception:
                pass

        def clear(_e=None):
            self.session.clear_log()
            self.log_lines.clear()
            refresh()

        # прокрутка к последнему событию после отрисовки
        try:
            self.page.run_task(self._scroll_log_end, list_view)
        except Exception:
            pass

        return ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Журнал", size=20, weight=ft.FontWeight.BOLD),
                        status,
                        ft.OutlinedButton(
                            "Обновить",
                            tooltip="Перечитать analysis.log (не более 500 строк)",
                            on_click=refresh,
                        ),
                        ft.OutlinedButton(
                            "Очистить",
                            tooltip="Очистить файл журнала",
                            on_click=clear,
                        ),
                        ft.OutlinedButton(
                            "Открыть файл",
                            tooltip="Открыть analysis.log во внешней программе",
                            on_click=lambda e: self.session.open_path(
                                str(self.session.app_dir / "analysis.log")
                            ),
                        ),
                    ],
                    wrap=True,
                ),
                list_view,
            ],
            expand=True,
            spacing=8,
        )

    async def _scroll_log_end(self, list_view: ft.ListView) -> None:
        import asyncio

        await asyncio.sleep(0.15)
        try:
            list_view.scroll_to(offset=-1, duration=200)
            self.page.update()
        except Exception:
            pass


def main(page: ft.Page | None = None) -> None:
    def _run(page: ft.Page) -> None:
        AnalizApp(page)

    if page is not None:
        _run(page)
        return
    assets = None
    for cand in (APP_ROOT / "assets", APP_ROOT / "_internal" / "assets"):
        if cand.is_dir():
            assets = str(cand)
            break
    ft.app(target=_run, assets_dir=assets)


if __name__ == "__main__":
    main()
