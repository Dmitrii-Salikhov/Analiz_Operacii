# Сборка Windows: папка dist/AnalizOperacii/ (onedir), не один файл.
# Запуск: pwsh ./build_windows.ps1
# Точка входа — Tk UI (app_desktop.py). Flet: отдельно python run_flet.py.

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$name = "AnalizOperacii"
$entry = "app_desktop.py"

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
    "--hidden-import", "yaml",
    "--hidden-import", "openpyxl",
    "--hidden-import", "pandas",
    "--hidden-import", "tkcalendar",
    "--hidden-import", "babel.numbers",
    "--collect-all", "tkcalendar",
    "--collect-all", "babel"
)

foreach ($d in $addData) {
    $args += @("--add-data", $d)
}
$args += $entry

Write-Host "PyInstaller $($args -join ' ')"
pyinstaller @args

# Дублируем служебные файлы в корень папки (удобно читать VERSION рядом с exe)
$out = "dist\$name"
Copy-Item -Force "VERSION" "$out\VERSION"
Copy-Item -Force "config.yaml" "$out\config.yaml"
Copy-Item -Force "RELEASE_NOTES.md" "$out\RELEASE_NOTES.md"
Copy-Item -Force "form14_overrides.yaml" "$out\form14_overrides.yaml"
Copy-Item -Force "requirements.txt" "$out\requirements.txt"
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

Write-Host "Готово: $out\$name.exe"
