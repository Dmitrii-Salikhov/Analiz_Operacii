"""Тесты сравнения версий updater (без сети)."""
from __future__ import annotations

import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))

from analyzers.updater import is_newer, parse_sha256_text, parse_version, read_local_version


def test_parse_version():
    assert parse_version("1.0.1") == (1, 0, 1)
    assert parse_version("v1.0.1") == (1, 0, 1)
    assert parse_version("1.0") == (1, 0)


def test_is_newer():
    assert is_newer("1.0.2", "1.0.1")
    assert not is_newer("1.0.1", "1.0.1")
    assert not is_newer("1.0.0", "1.0.1")
    assert is_newer("2.0.0", "1.9.9")


def test_local_version_file():
    ver = read_local_version(APP)
    assert ver
    assert parse_version(ver) >= (1, 0, 0)


def test_parse_sha256_text():
    h = "a" * 64
    assert parse_sha256_text(f"{h}  AnalizOperacii-Windows.zip") == h
    assert parse_sha256_text(h) == h


if __name__ == "__main__":
    test_parse_version()
    test_is_newer()
    test_local_version_file()
    test_parse_sha256_text()
    print("OK")
