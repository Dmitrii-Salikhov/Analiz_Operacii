# ui_flet/screens/form14.py
"""Конструктор ФСН 14 — схемный UI поверх form14_* backend (Flet 0.86+)."""
from __future__ import annotations

from typing import Callable, List, Optional, Set

import flet as ft

from analyzers.category_registry import save_config
from analyzers.dept_config import sync_histology_overrides_to_config
from analyzers.form14_export import (
    DEFAULT_XLSX,
    build_mapping_rows,
    form14_line_choices,
    parse_line_choice,
    write_form14_excel,
)
from analyzers.form14_map import FORM14_LINES
from analyzers.form14_overrides import (
    clear_override,
    default_path,
    import_overrides_from_excel,
    load_overrides,
    merge_store,
    save_overrides,
    set_histology,
    set_override,
)
from ui_flet.app_context import AppContext
from ui_flet.schema_loader import load_schema

CONF_RU = {
    "high": "высокая",
    "medium": "средняя",
    "low": "низкая",
}


def _ru_confidence(raw: str) -> str:
    return CONF_RU.get(str(raw or "").strip().lower(), str(raw or ""))


def _ru_rule(raw: str) -> str:
    """Переводит служебные пометки правила в понятный русский текст."""
    t = str(raw or "").strip()
    if not t:
        return ""
    repl = [
        ("override YAML", "ручная правка (файл)"),
        ("override", "ручная правка"),
        ("MANUAL_LINE_BY_CODE", "калибровка по коду"),
        ("MANUAL_LINE_BY_CATEGORY", "калибровка по категории"),
        ("MANUAL_LINE", "калибровка"),
        ("keyword/class:", "по словам и классу:"),
        ("keyword:", "по словам:"),
        ("class+keyword:", "класс и слова:"),
        ("class:", "по классу:"),
        ("default:", "по умолчанию:"),
        ("fallback", "запасной вариант"),
        ("A16", "код A16"),
    ]
    out = t
    for a, b in repl:
        out = out.replace(a, b)
    return out


class Form14ConstructorView:
    def __init__(
        self,
        page: ft.Page,
        ctx: AppContext,
        on_back: Optional[Callable] = None,
        *,
        default_dept_key: Optional[str] = None,
    ):
        self.page = page
        self.ctx = ctx
        self.on_back = on_back
        self.schema = load_schema("form14_constructor.schema.yaml")
        self.ov_path = default_path(ctx.app_dir)
        self.store = load_overrides(self.ov_path)
        self.dirty = False
        self.rows: List[dict] = []
        self.checked: Set[int] = set()
        self._visible: List[dict] = []

        self.status = ft.Text("")
        self.search = ft.TextField(
            label="Поиск",
            width=220,
            dense=True,
            on_change=lambda e: self.refresh(),
            tooltip="Поиск по коду A16 или названию операции",
        )
        # для не-ЛОР по умолчанию показываем все строки отделения, не только спорные
        prefer_key = default_dept_key or ctx.current_summary_key()
        only_disputed_default = prefer_key in (None, "", "lor", "все")
        self.only_disp = ft.Switch(
            label="Только спорные (низкая уверенность / стр. 21)",
            value=only_disputed_default,
            on_change=lambda e: self.refresh(),
            tooltip="Показать коды с низкой уверенностью или отнесённые к стр. 21 «прочие»",
        )
        dept_opts = [
            ft.dropdown.Option(key="все", text="Все отделения"),
            *[
                ft.dropdown.Option(key=k, text=ctx.dept_full_name(k))
                for k in ctx.dept_keys()
            ],
        ]
        opt_keys = {o.key for o in dept_opts}
        self.dept = ft.Dropdown(
            label="Отделение",
            width=340,
            dense=True,
            options=dept_opts,
            value=prefer_key if prefer_key in opt_keys else "все",
            on_select=self._on_dept_filter,
            tooltip="Фильтр: операции и коды выбранного отделения",
        )
        line_filter_opts = [ft.dropdown.Option(key="все", text="Все строки формы")]
        for k, name in FORM14_LINES.items():
            if k == "1":
                continue
            line_filter_opts.append(
                ft.dropdown.Option(key=str(k), text=f"{k} — {name}")
            )
        self.line_filter = ft.Dropdown(
            label="Фильтр: строка формы",
            width=340,
            dense=True,
            options=line_filter_opts,
            value="все",
            on_select=lambda e: self.refresh(),
            tooltip="Показать только коды с выбранной итоговой строкой формы № 14",
        )
        choices = form14_line_choices()
        self.line_dd = ft.Dropdown(
            label="Строка формы № 14",
            width=440,
            dense=True,
            options=[ft.dropdown.Option(c) for c in choices],
            value=choices[0] if choices else None,
            tooltip="Целевая строка таблиц 4000/4001 формы № 14",
        )
        self.comment = ft.TextField(
            label="Комментарий",
            width=240,
            dense=True,
            tooltip="Пояснение к ручному назначению строки",
        )
        self.morph_sw = ft.Switch(
            label="Морфология",
            value=False,
            tooltip="Морфологическое исследование (гистология) для отмеченных кодов",
        )
        self.list_col = ft.Column(spacing=0, scroll=ft.ScrollMode.AUTO, expand=True)

        self.file_picker = ft.FilePicker()
        try:
            page.services.append(self.file_picker)
        except Exception:
            page.overlay.append(self.file_picker)

    def _on_dept_filter(self, _e=None) -> None:
        # при смене отделения сбрасываем «только спорные» для не-ЛОР — иначе список пустой
        key = self.dept.value or "все"
        if key not in ("lor", "все") and self.only_disp.value:
            self.only_disp.value = False
        self.refresh()

    def build(self) -> ft.Control:
        toolbar = ft.Row(
            [
                *(
                    [
                        ft.OutlinedButton(
                            "← Назад",
                            tooltip="Вернуться к предыдущему экрану",
                            on_click=lambda e: self.on_back() if self.on_back else None,
                        )
                    ]
                    if self.on_back
                    else []
                ),
                ft.Text("Конструктор формы № 14", size=20, weight=ft.FontWeight.BOLD),
                ft.Container(expand=True),
                self.status,
            ]
        )
        filters = ft.Row(
            [self.dept, self.line_filter, self.only_disp, self.search], wrap=True
        )
        edit = ft.Row(
            [
                self.line_dd,
                self.comment,
                ft.FilledButton(
                    "Назначить",
                    tooltip="Присвоить выбранным кодам строку формы № 14",
                    on_click=self._assign,
                ),
                ft.OutlinedButton(
                    "Отвязать",
                    tooltip="Вернуть авто-строку формы № 14 у отмеченных (морфологию не сбрасывает)",
                    on_click=self._reset,
                ),
                self.morph_sw,
                ft.OutlinedButton(
                    "Применить морфологию",
                    tooltip="Задать флаг морфологии для отмеченных кодов",
                    on_click=self._apply_morph,
                ),
                ft.FilledButton(
                    "Сохранить",
                    tooltip="Записать правки в form14_overrides.yaml и синхронизировать морфологию в config",
                    on_click=self._save,
                ),
                ft.OutlinedButton(
                    "Экспорт Excel…",
                    tooltip="Выгрузить таблицу соответствия в Excel",
                    on_click=self._export_ask,
                ),
                ft.OutlinedButton(
                    "Импорт Excel…",
                    tooltip="Загрузить ручные строки из Excel-калибровки",
                    on_click=self._import_ask,
                ),
            ],
            wrap=True,
        )
        hint = ft.Text(
            "Приоритет: ручная правка → для ЛОР калибровка сводной → авто по коду A16 и словам "
            "(для каждого отделения — свои категории). «Отвязать» возвращает авто-строку; "
            "морфология пишется в overrides и при сохранении — в категории config.",
            size=12,
            color=ft.Colors.ON_SURFACE_VARIANT,
        )
        header = ft.Container(
            content=ft.Row(
                [
                    ft.Text("", width=40),
                    ft.Text("Код", width=120, weight=ft.FontWeight.BOLD),
                    ft.Text("Название", expand=True, weight=ft.FontWeight.BOLD),
                    ft.Text("Отделение", width=200, weight=ft.FontWeight.BOLD),
                    ft.Text("Авто", width=45, weight=ft.FontWeight.BOLD),
                    ft.Text("Итог", width=45, weight=ft.FontWeight.BOLD),
                    ft.Text("Морф.", width=50, weight=ft.FontWeight.BOLD),
                    ft.Text("Уверенность", width=85, weight=ft.FontWeight.BOLD),
                    ft.Text("Основание", width=180, weight=ft.FontWeight.BOLD),
                ]
            ),
            padding=ft.Padding(left=8, right=8, top=4, bottom=4),
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST,
        )
        root = ft.Column(
            [toolbar, filters, edit, hint, header, self.list_col],
            expand=True,
            spacing=8,
        )
        self.refresh()
        return root

    def refresh(self, *_a) -> None:
        keys = None if self.dept.value == "все" else [self.dept.value]
        self.rows = build_mapping_rows(self.ctx.config, overrides=self.store, dept_keys=keys)
        q = (self.search.value or "").strip().lower()
        line_f = str(self.line_filter.value or "все").strip()
        self.list_col.controls.clear()
        self._visible = []
        self.checked.clear()
        for r in self.rows:
            final = str(r.get("Строка_ФСН14") or "")
            conf = str(r.get("Уверенность") or "")
            if line_f and line_f != "все" and final != line_f:
                continue
            if self.only_disp.value and not (conf == "low" or final == "21"):
                if not r.get("Строка_ФСН14_ручная"):
                    continue
            code = str(r.get("Код") or "")
            name = str(r.get("Наименование_КСГ") or r.get("Категория") or "")
            if q and q not in code.lower() and q not in name.lower():
                continue
            idx = len(self._visible)
            self._visible.append(r)
            cb = ft.Checkbox(value=False, data=idx, on_change=self._toggle)
            conf_ru = _ru_confidence(conf)
            rule_ru = _ru_rule(str(r.get("Правило") or ""))
            morph = "да" if r.get("histology") else "нет"
            self.list_col.controls.append(
                ft.Container(
                    content=ft.Row(
                        [
                            cb,
                            ft.Text(code, width=120, size=12),
                            ft.Text((name or "")[:50], expand=True, size=12),
                            ft.Text(
                                str(
                                    r.get("Отделение")
                                    or self.ctx.dept_full_name(str(r.get("summary_key") or ""))
                                ),
                                width=200,
                                size=12,
                            ),
                            ft.Text(str(r.get("Авто") or ""), width=45, size=12),
                            ft.Text(final, width=45, size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(morph, width=50, size=12),
                            ft.Text(conf_ru, width=85, size=12),
                            ft.Text(rule_ru[:32], width=180, size=11),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    padding=ft.Padding(left=4, right=4, top=2, bottom=2),
                    border=ft.Border(
                        bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)
                    ),
                )
            )
        nov = len(self.store.get("by_code") or {}) + len(self.store.get("by_category") or {})
        star = " *" if self.dirty else ""
        self.status.value = f"Показано: {len(self._visible)} | ручных правок: {nov}{star}"
        self.page.update()

    def _toggle(self, e: ft.ControlEvent) -> None:
        idx = int(e.control.data)
        if e.control.value:
            self.checked.add(idx)
            r = self._visible[idx]
            line = str(r.get("Строка_ФСН14_ручная") or r.get("Строка_ФСН14") or "")
            for opt in self.line_dd.options or []:
                text = str(opt.key or "")
                if text.startswith(line + " —") or text.startswith(line + " -") or text == line:
                    self.line_dd.value = text
                    break
            self.comment.value = str(r.get("Комментарий") or "")
            self.morph_sw.value = bool(r.get("histology"))
            self.page.update()
        else:
            self.checked.discard(idx)

    def _targets(self) -> List[dict]:
        return [self._visible[i] for i in sorted(self.checked) if i < len(self._visible)]

    def _assign(self, _e=None) -> None:
        targets = self._targets()
        if not targets:
            self._snack("Отметьте строки галочкой")
            return
        line = parse_line_choice(self.line_dd.value or "")
        if not line:
            self._snack("Выберите строку формы № 14")
            return
        comment = (self.comment.value or "").strip()
        for r in targets:
            code = str(r.get("Код") or "").strip()
            cat = str(r.get("Категория") or "").strip()
            self.store = set_override(
                self.store,
                line=line,
                code=code,
                category="" if code else cat,
                comment=comment,
                by="конструктор",
            )
        self.dirty = True
        self.refresh()
        self._snack(f"Назначено → стр. {line} ({len(targets)})")

    def _apply_morph(self, _e=None) -> None:
        targets = self._targets()
        if not targets:
            self._snack("Отметьте строки галочкой")
            return
        value = bool(self.morph_sw.value)
        for r in targets:
            code = str(r.get("Код") or "").strip()
            cat = str(r.get("Категория") or "").strip()
            self.store = set_histology(
                self.store,
                value=value,
                code=code,
                category="" if code else cat,
                by="конструктор",
            )
        self.dirty = True
        self.refresh()
        self._snack(f"Морфология → {'да' if value else 'нет'} ({len(targets)})")

    def _reset(self, _e=None) -> None:
        targets = self._targets()
        if not targets:
            self._snack("Отметьте строки галочкой")
            return
        for r in targets:
            self.store = clear_override(
                self.store,
                code=str(r.get("Код") or "").strip(),
                category=str(r.get("Категория") or "").strip(),
            )
        self.dirty = True
        self.refresh()
        self._snack(f"Отвязано от строки формы ({len(targets)})")

    def _save(self, _e=None) -> None:
        save_overrides(self.store, self.ov_path)
        n_hist = sync_histology_overrides_to_config(self.ctx.config, self.store)
        if n_hist:
            save_config(self.ctx.config, self.ctx.app_dir / "config.yaml")
        self.dirty = False
        self.refresh()
        extra = f", морфология в config: {n_hist}" if n_hist else ""
        self._snack(f"Сохранено: {self.ov_path.name}{extra}")

    async def _export_ask(self, _e=None) -> None:
        path = await self.file_picker.save_file(
            dialog_title="Экспорт соответствия форме № 14",
            file_name=DEFAULT_XLSX,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            initial_directory=str(self.ctx.app_dir),
        )
        if not path:
            path = str(self.ctx.app_dir / DEFAULT_XLSX)
        write_form14_excel(path, self.ctx.config, overrides=self.store)
        self._snack(f"Экспорт: {path}")

    async def _import_ask(self, _e=None) -> None:
        files = await self.file_picker.pick_files(
            dialog_title="Импорт калибровки формы № 14",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            allow_multiple=False,
            initial_directory=str(self.ctx.app_dir),
        )
        if not files:
            self._snack("Файл не выбран")
            return
        path = getattr(files[0], "path", None)
        if not path:
            self._snack("Не удалось получить путь к файлу")
            return
        frag, n = import_overrides_from_excel(path)
        if n == 0:
            self._snack("Нет заполненных «Строка_ФСН14_ручная» / «Морфология»")
            return
        self.store = merge_store(self.store, frag)
        self.dirty = True
        self.refresh()
        self._snack(f"Импорт: {n} — нажмите «Сохранить»")

    def _snack(self, msg: str) -> None:
        try:
            self.page.show_dialog(ft.SnackBar(content=ft.Text(msg)))
        except Exception:
            self.status.value = msg
            self.page.update()
