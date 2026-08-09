# Сборка Windows: папка dist/AnalizOperacii/ (onedir), не один файл.
# Точка входа — Flet UI (run_flet.py). Tk app_desktop.py остаётся для разработки.
# Запуск: pwsh ./build_windows.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$name = "AnalizOperacii"
$entry = "run_flet.py"

# Прогрев desktop-клиента Flet (скачивает client при первом импорте)
Write-Host "Проверка Flet / flet_desktop…"
python -c "import flet; import flet_desktop; print('flet', getattr(flet,'__version__', '?'))"

$addData = @(
    "VERSION;.",
    "config.yaml;.",
    "requirements.txt;.",
    "RELEASE_NOTES.md;.",
    "form14_overrides.yaml;.",
    "schemas;schemas"
)

if (Test-Path "KSGoperacii.csv") {
    $addData += "KSGoperacii.csv;."
}
if (Test-Path "Операции сводная 2026.xlsx") {
    $addData += "Операции сводная 2026.xlsx;."
}

$args = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--name", $name,
    "--collect-submodules", "analyzers",
    "--collect-submodules", "ui_flet",
    "--collect-all", "flet",
    "--collect-all", "flet_desktop",
    "--hidden-import", "flet",
    "--hidden-import", "flet_desktop",
    "--hidden-import", "yaml",
    "--hidden-import", "openpyxl",
    "--hidden-import", "pandas"
)

foreach ($d in $addData) {
    $args += @("--add-data", $d)
}
$args += $entry

Write-Host "PyInstaller $($args -join ' ')"
pyinstaller @args

# Дублируем служебные файлы в корень папки (рядом с AnalizOperacii.exe)
$out = "dist\$name"
Copy-Item -Force "VERSION" "$out\VERSION"
Copy-Item -Force "config.yaml" "$out\config.yaml"
Copy-Item -Force "RELEASE_NOTES.md" "$out\RELEASE_NOTES.md"
Copy-Item -Force "form14_overrides.yaml" "$out\form14_overrides.yaml"
Copy-Item -Force "requirements.txt" "$out\requirements.txt"
Copy-Item -Force "run_flet.py" "$out\run_flet.py"
if (Test-Path "schemas") {
    if (Test-Path "$out\schemas") { Remove-Item -Recurse -Force "$out\schemas" }
    Copy-Item -Recurse -Force "schemas" "$out\schemas"
}
if (Test-Path "KSGoperacii.csv") {
    Copy-Item -Force "KSGoperacii.csv" "$out\KSGoperacii.csv"
}
if (Test-Path "Операции сводная 2026.xlsx") {
    Copy-Item -Force "Операции сводная 2026.xlsx" "$out\Операции сводная 2026.xlsx"
}

$exe = Join-Path $out "$name.exe"
if (-not (Test-Path $exe)) {
    throw "Нет $exe — сборка Flet не создала exe"
}
Write-Host "Готово (Flet): $exe"
