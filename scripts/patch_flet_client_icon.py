#!/usr/bin/env python3
"""Подготовить Windows-клиент Flet с нашей иконкой (только на Windows CI/сборке).

PyInstaller --icon меняет только AnalizOperacii.exe. Окно и панель задач —
это flet.exe; его нужно патчить так же, как делает `flet pack --icon`.
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


def main() -> int:
    if sys.platform != "win32":
        print("skip: patch_flet_client_icon только для Windows", file=sys.stderr)
        return 2

    parser = argparse.ArgumentParser()
    parser.add_argument("icon", type=Path, help="Путь к .ico")
    parser.add_argument(
        "out_dir",
        type=Path,
        help="Каталог вывода (внутри появится flet/flet.exe)",
    )
    args = parser.parse_args()

    icon = args.icon.resolve()
    if not icon.is_file():
        print(f"нет иконки: {icon}", file=sys.stderr)
        return 1

    from flet_cli.__pyinstaller.utils import copy_flet_bin
    from flet_cli.__pyinstaller.win_utils import update_flet_view_icon

    temp = copy_flet_bin()
    if not temp:
        print("не удалось скопировать клиент Flet", file=sys.stderr)
        return 1

    exe = Path(temp) / "flet" / "flet.exe"
    if not exe.is_file():
        print(f"нет {exe}", file=sys.stderr)
        shutil.rmtree(temp, ignore_errors=True)
        return 1

    print(f"Патч иконки: {exe} <- {icon}")
    update_flet_view_icon(str(exe), str(icon))

    out = args.out_dir.resolve()
    dest = out / "flet"
    if dest.exists():
        shutil.rmtree(dest)
    out.mkdir(parents=True, exist_ok=True)
    shutil.copytree(Path(temp) / "flet", dest)
    shutil.rmtree(temp, ignore_errors=True)
    print(f"Готово: {dest / 'flet.exe'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
