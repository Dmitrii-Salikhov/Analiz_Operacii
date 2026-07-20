# Сборка Windows: папка dist/AnalizOperacii/ (onedir), не один файл.
# Запуск: pwsh ./build_windows.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$name = "AnalizOperacii"
$entry = "app_desktop.py"

$addData = @(
    "VERSION;.",
    "config.yaml;.",
    "requirements.txt;."
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
    "--hidden-import", "analyzers",
    "--hidden-import", "analyzers.updater",
    "--hidden-import", "analyzers.surgery",
    "--hidden-import", "analyzers.summary_writer",
    "--hidden-import", "analyzers.summary_layout",
    "--hidden-import", "analyzers.category_registry",
    "--hidden-import", "analyzers.release_notes",
    "--hidden-import", "analyzers.form_4001",
    "--hidden-import", "analyzers.io_utils",
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
if (Test-Path "KSGoperacii.csv") {
    Copy-Item -Force "KSGoperacii.csv" "$out\KSGoperacii.csv"
}
if (Test-Path "Операции сводная 2026.xlsx") {
    Copy-Item -Force "Операции сводная 2026.xlsx" "$out\Операции сводная 2026.xlsx"
}

Write-Host "Готово: $out\$name.exe"
