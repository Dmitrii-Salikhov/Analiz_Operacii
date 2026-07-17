#!/bin/bash
cd "$(dirname "$0")"
exec "$(dirname "$0")/.venv/bin/python" app_desktop.py
