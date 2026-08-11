# Сборка Windows: папка dist/AnalizOperacii/ (onedir), не один файл.
# Точка входа — Flet UI (run_flet.py). Tk app_desktop.py остаётся для разработки.
# Запуск: pwsh ./build_windows.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$name = "AnalizOperacii"
$entry = "run_flet.py"
$icon = "assets\app_icon.ico"
$fletViewOut = "build\flet_view"

# Прогрев desktop-клиента Flet (скачивает client при первом импорте)
Write-Host "Check Flet / flet_desktop..."
python -c "import flet; import flet_desktop; print('flet', getattr(flet,'__version__', '?'))"

if (-not (Test-Path $icon)) {
    throw "Missing icon: $icon"
}

# Patch flet.exe (window/taskbar); PyInstaller --icon alone is not enough
Write-Host "Patch Flet client icon..."
if (Test-Path $fletViewOut) { Remove-Item -Recurse -Force $fletViewOut }
python scripts\patch_flet_client_icon.py $icon $fletViewOut
if ($LASTEXITCODE -ne 0) {
    throw "patch_flet_client_icon.py failed with exit $LASTEXITCODE"
}
if (-not (Test-Path "$fletViewOut\flet\flet.exe")) {
    throw "Missing patched client: $fletViewOut\flet\flet.exe"
}

$addData = @(
    "VERSION;.",
    "config.yaml;.",
    "requirements.txt;.",
    "RELEASE_NOTES.md;.",
    "form14_overrides.yaml;.",
    "schemas;schemas",
    "assets;assets",
    "$fletViewOut;flet_view"
)

if (Test-Path "KSGoperacii.csv") {
    $addData += "KSGoperacii.csv;."
}
if (Test-Path "Операции сводная 2026.xlsx") {
    $addData += "Операции сводная 2026.xlsx;."
}

# Не использовать имя $args — в PowerShell это автоматическая переменная
$pyiArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--name", $name,
    "--icon", $icon,
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
    $pyiArgs += @("--add-data", $d)
}
$pyiArgs += $entry

Write-Host "Exe icon: $icon"
Write-Host "PyInstaller $($pyiArgs -join ' ')"
pyinstaller @pyiArgs

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
if (Test-Path "assets") {
    if (Test-Path "$out\assets") { Remove-Item -Recurse -Force "$out\assets" }
    Copy-Item -Recurse -Force "assets" "$out\assets"
}
# Клиент с иконкой рядом с exe (FLET_VIEW_PATH в run_flet.py)
if (Test-Path $fletViewOut) {
    if (Test-Path "$out\flet_view") { Remove-Item -Recurse -Force "$out\flet_view" }
    Copy-Item -Recurse -Force $fletViewOut "$out\flet_view"
}
if (Test-Path "KSGoperacii.csv") {
    Copy-Item -Force "KSGoperacii.csv" "$out\KSGoperacii.csv"
}
if (Test-Path "Операции сводная 2026.xlsx") {
    Copy-Item -Force "Операции сводная 2026.xlsx" "$out\Операции сводная 2026.xlsx"
}

$exe = Join-Path $out "$name.exe"
if (-not (Test-Path $exe)) {
    throw "Missing $exe — Flet build did not produce exe"
}
Write-Host "Done (Flet): $exe"
