# ui_flet/__init__.py
"""Flet UI поверх analyzers (схемные параметры)."""
__all__ = ["main"]


def main():
    from ui_flet.app import main as _main

    _main()
