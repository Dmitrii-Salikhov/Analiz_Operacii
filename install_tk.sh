#!/bin/bash
set -e
cd "$(dirname "$0")"

if ! command -v brew >/dev/null 2>&1; then
  echo "Homebrew не найден. Установка:"
  echo '/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
  echo "Если curl таймаут — поставьте Python с сайта python.org (включает Tcl/Tk 8.6)."
  exit 1
fi

brew install python-tk@3.12
PY="$(brew --prefix python-tk@3.12)/bin/python3.12"
if [ ! -x "$PY" ]; then
  PY="$(brew --prefix python@3.12)/bin/python3.12"
fi

"$PY" -m venv .venv-tk
source .venv-tk/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
python -c "import tkinter as tk; r=tk.Tk(); print('Tk', r.tk.call('info','patchlevel')); r.destroy()"
echo "Запуск: source .venv-tk/bin/activate && python app_desktop.py"
