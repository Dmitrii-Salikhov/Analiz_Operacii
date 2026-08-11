# Инвентарь UI и параметров (фиксировано для миграции на Flet)

Дата фиксации: 2026-07-26.  
Источник истины домена: пакет `analyzers/`. UI (Tk или Flet) только вызывает API, правила не дублирует.

Статусы в Flet: `done` — реализовано в Flet; `shell` — пункт в навигации + вызов backend/заглушка с тем же контрактом; `backend` — только Python API, UI позже.

---

## 1. Экраны / вкладки

| ID | Название | Статус |
|----|----------|--------|
| work | Работа (источники, настройки) | done |
| preview | Превью: категории / итоги / 4001 | done |
| emk | Расхождения ЭМК | done |
| uncl | Не классифицировано | done |
| disp | Спорные | done |
| form14 | Конструктор ФСН 14 | done |
| log | Журнал | done |

---

## 2. Меню и действия (нельзя потерять)

| ID | Действие | Backend API | Статус |
|----|----------|-------------|--------|
| load_surg | Опержурнал(ы)… | `io_utils.read_table`, `OperationsStore.add`, `SurgeryAnalyzer` | shell |
| load_surg_folder | Опержурналы из папки… | то же | shell |
| load_emk | ЭМК… | `emk_loader.read_emk_stationary_report` | shell |
| choose_summary | Сводная… | путь + `ui_settings` | shell |
| write_excel | Записать в Excel… | `SummaryWriter.write`, `file_lock`, `write_verify`, backup | shell |
| open_excel | Открыть Excel | OS open | shell |
| restore_backup | Восстановить из бэкапа… | `backup_utils` | shell |
| export_simple | Экспорт простого отчёта… | `export_report.export_month_like_summary` | shell |
| create_year | Создать сводную на год… | `year_template.create_year_summary` | shell |
| export_uncl | Экспорт неклассифицированных… | DataFrame → csv/xlsx | shell |
| export_problems | Экспорт проблемных кодов… | `problem_codes` | shell |
| inventory | Инвентаризация отделения… | `dept_inventory` | shell |
| create_dept_summary | Создать сводную для отделения… | `dept_template.create_from_summary_cfg` | shell |
| classify_emk_kind | План/экстр по ЭМК… | `emk_kind_classify` + `save_config` | shell |
| form14_ctor | Конструктор ФСН 14… | `form14_*` | done |
| add_category | Добавить операцию… | `category_registry` ± `summary_layout` | shell |
| delete_category | Удалить операцию… | `unregister_category` ± delete row | shell |
| clear_store | Очистить | `OperationsStore.clear` | shell |
| check_updates | Проверить обновления… | `updater` | done |
| whats_new | Что нового… | `release_notes` | shell |
| about | О программе | VERSION | shell |
| refresh_preview | Обновить превью | `build_summary_tables` | shell |
| theme_toggle | Тема light/dark | `ui_settings.theme` | shell |
| edit_keywords | Править ключи… | `update_category_keywords_file` | shell |
| assign_dispute | Назначить категорию (спорные) | ручная метка в ops | shell |

Горячие клавиши: Ctrl/Cmd+O load_surg, Ctrl/Cmd+S write_excel, Ctrl/Cmd+C copy selection.

---

## 3. Параметры конфигурации

### 3.1. `config.yaml`
- `departments.main`, `departments.list[]`
- `thresholds.max_bed_days`, `thresholds.pension_age`
- `updates.*` (enabled, github_*, check_on_startup, check_interval_days, zip/sha256)
- `department_profiles.{dept}.{summary_key, rubricator}`
- `summaries.{lor,surg1,surg2,pedsurg,traum}` и legacy `summary`
- `surgery_categories_by_dept.*` и legacy `surgery_categories`

### 3.2. SummaryCfg
`default_path`, `backup_keep`, `year`, `sheet_names`, `category_rows`, `totals_rows`, `plan_categories`, `emergency_categories`, `form_4001` (только lor)

### 3.3. `form_4001` (ЛОР)
`enabled`, `line_rows`, `parent_row`, `total_row`, `pension_row`, `cols.*`, `line_categories`, `hist_categories`, `endo_categories`

### 3.4. Категория
`category`, `codes[]`, `group`, `line`, `histology`, `name_keywords[]`

### 3.5. `ui_settings.json`
`summary_path`, `department`, `year`, `hide_zeros`, `filter_enabled`, `start_date`, `end_date`, `write_weeks`, `write_form`, `plan_mode`, `last_surg_dir`, `last_emk_dir`, `last_update_check`, `last_seen_version`, `summary_paths_by_dept`, `theme`

### 3.6. `form14_overrides.yaml`
`by_code` / `by_category` → `{line, comment, by, at}`

---

## 4. Приоритеты правил (не менять)

**ФСН 14:** override YAML (код > категория) → `MANUAL_LINE_BY_CATEGORY` → `MANUAL_LINE_BY_CODE` → класс A16 + keywords → стр. 21.

**План/экстр:** `is_forced_emergency_category` > ЭМК-статистика > списки config (forced всё равно применяется при apply).

Forced emergency: «дренирован»+«абсцесс» ИЛИ «вскрытие»+(«флегмон»|«абсцесс»).

---

## 5. Файлы данных

`config.yaml`, `ui_settings.json`, `form14_overrides.yaml`, `VERSION`, `RELEASE_NOTES.md`, `analysis.log`, `KSGoperacii.csv`, сводные xlsx, `маппинг_ФСН14_4000_4001.xlsx`, `backups/`, `Отчеты других отделений/`.

---

## 6. Checklist миграции

См. раздел «Что нельзя потерять» — все пункты из инвентаризации должны либо быть `done` в Flet, либо оставаться доступны через Tk (`app_desktop.py`) до порта.

Tk (`app_desktop.py`) сохраняется как fallback, пока Flet-покрытие неполное.
