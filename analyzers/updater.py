# analyzers/updater.py
"""Обновление приложения с GitHub Releases (ZIP папки Windows + SHA-256)."""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import ssl
import tempfile
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

ZIP_FILENAME = "AnalizOperacii-Windows.zip"
SHA256_FILENAME = "AnalizOperacii-Windows.zip.sha256"
EXE_NAME = "AnalizOperacii.exe"

SKIP_DIR_NAMES = {
    ".git",
    ".venv",
    ".venv_py39_bak",
    "__pycache__",
    ".cursor",
    "node_modules",
    ".update_backup_",
}
SKIP_FILE_SUFFIXES = {".xlsx", ".xls", ".csv", ".bak", ".log"}
SKIP_FILE_NAMES = {
    "ui_settings.json",
    "analysis.log",
    ".ds_store",
}


@dataclass
class UpdateInfo:
    local_version: str
    remote_version: str
    tag: str
    name: str
    body: str
    zip_url: str
    html_url: str
    source: str  # "release-asset" | "release-zipball" | "branch"
    sha256_url: Optional[str] = None


def read_local_version(app_dir: Path) -> str:
    path = Path(app_dir) / "VERSION"
    if path.exists():
        text = path.read_text(encoding="utf-8").strip().splitlines()
        if text:
            return text[0].strip().lstrip("vV")
    return "0.0.0"


def parse_version(ver: str) -> Tuple[int, ...]:
    ver = (ver or "0").strip().lstrip("vV")
    parts: List[int] = []
    for chunk in re.split(r"[^\d]+", ver):
        if chunk.isdigit():
            parts.append(int(chunk))
    return tuple(parts) if parts else (0,)


def is_newer(remote: str, local: str) -> bool:
    return parse_version(remote) > parse_version(local)


def _ssl_context() -> ssl.SSLContext:
    try:
        import certifi  # type: ignore

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _github_headers(token: Optional[str] = None) -> Dict[str, str]:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "Analiz-Operacii-Updater",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def _http_json(url: str, token: Optional[str] = None, timeout: int = 30) -> Any:
    req = urllib.request.Request(url, headers=_github_headers(token))
    with urllib.request.urlopen(req, context=_ssl_context(), timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _http_download(url: str, dest: Path, token: Optional[str] = None, timeout: int = 180) -> None:
    headers = _github_headers(token)
    headers = dict(headers)
    headers["Accept"] = "application/octet-stream"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=_ssl_context(), timeout=timeout) as resp:
        dest.write_bytes(resp.read())


def resolve_token(cfg: Dict[str, Any]) -> Optional[str]:
    env_name = str(cfg.get("github_token_env") or "GITHUB_TOKEN")
    token = os.environ.get(env_name) or os.environ.get("GH_TOKEN")
    if token:
        return token.strip()
    raw = cfg.get("github_token")
    if raw:
        return str(raw).strip()
    return None


def find_release_asset(release: dict, filename: str) -> Optional[dict]:
    for asset in release.get("assets") or []:
        if asset.get("name") == filename:
            return asset
    return None


def parse_sha256_text(text: str, expected_filename: str = ZIP_FILENAME) -> Optional[str]:
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        match = re.match(r"^([A-Fa-f0-9]{64})(?:\s+\*?(\S+))?$", line)
        if not match:
            continue
        digest, name = match.group(1), match.group(2)
        if name is None or os.path.basename(name) == expected_filename:
            return digest.lower()
    return None


def compute_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def check_for_update(app_dir: Path, cfg: Dict[str, Any]) -> Optional[UpdateInfo]:
    if not cfg or not cfg.get("enabled", True):
        return None

    owner = str(cfg.get("github_owner") or "").strip()
    repo = str(cfg.get("github_repo") or "").strip()
    if not owner or not repo:
        raise ValueError(
            "В config.yaml укажите updates.github_owner и updates.github_repo"
        )

    branch = str(cfg.get("branch") or "main").strip()
    token = resolve_token(cfg)
    local = read_local_version(app_dir)
    zip_name = str(cfg.get("zip_filename") or ZIP_FILENAME)
    sha_name = str(cfg.get("sha256_filename") or SHA256_FILENAME)

    release_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    try:
        data = _http_json(release_url, token=token)
        tag = str(data.get("tag_name") or "").strip()
        remote = tag.lstrip("vV")
        if remote and is_newer(remote, local):
            zip_asset = find_release_asset(data, zip_name)
            sha_asset = find_release_asset(data, sha_name) or find_release_asset(
                data, "SHA256SUMS"
            )
            if zip_asset and zip_asset.get("browser_download_url"):
                return UpdateInfo(
                    local_version=local,
                    remote_version=remote,
                    tag=tag,
                    name=str(data.get("name") or tag),
                    body=str(data.get("body") or ""),
                    zip_url=str(zip_asset["browser_download_url"]),
                    sha256_url=(
                        str(sha_asset["browser_download_url"])
                        if sha_asset and sha_asset.get("browser_download_url")
                        else None
                    ),
                    html_url=str(
                        data.get("html_url")
                        or f"https://github.com/{owner}/{repo}/releases"
                    ),
                    source="release-asset",
                )
            # нет готового ZIP — исходники релиза
            zip_url = str(data.get("zipball_url") or "")
            if not zip_url:
                zip_url = f"https://api.github.com/repos/{owner}/{repo}/zipball/{tag}"
            return UpdateInfo(
                local_version=local,
                remote_version=remote,
                tag=tag,
                name=str(data.get("name") or tag),
                body=str(data.get("body") or ""),
                zip_url=zip_url,
                html_url=str(data.get("html_url") or ""),
                source="release-zipball",
            )
    except urllib.error.HTTPError as e:
        if e.code not in (404,):
            if e.code >= 500 or e.code in (401, 403):
                raise RuntimeError(f"GitHub API ({e.code}): {e.reason}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Нет сети или GitHub недоступен: {e.reason}") from e

    raw_url = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/VERSION"
    try:
        req = urllib.request.Request(
            raw_url, headers={"User-Agent": "Analiz-Operacii-Updater"}
        )
        with urllib.request.urlopen(req, context=_ssl_context(), timeout=30) as resp:
            remote = resp.read().decode("utf-8").strip().splitlines()[0].strip().lstrip("vV")
        if remote and is_newer(remote, local):
            zip_url = f"https://api.github.com/repos/{owner}/{repo}/zipball/{branch}"
            return UpdateInfo(
                local_version=local,
                remote_version=remote,
                tag=branch,
                name=f"{branch} @ {remote}",
                body="Обновление с ветки (файл VERSION).",
                zip_url=zip_url,
                html_url=f"https://github.com/{owner}/{repo}/tree/{branch}",
                source="branch",
            )
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise RuntimeError(
                f"Репозиторий или VERSION не найдены: {owner}/{repo} (ветка {branch}).\n"
                "Создайте репозиторий Analiz_Operacii на GitHub и запушьте код."
            ) from e
        raise RuntimeError(f"GitHub ({e.code}): {e.reason}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Нет сети или GitHub недоступен: {e.reason}") from e

    return None


def _should_skip_path(rel: Path, *, include_config: bool, preserve_data: bool) -> bool:
    parts_lower = [p.lower() for p in rel.parts]
    if any(p in SKIP_DIR_NAMES or p.startswith(".update_backup_") for p in parts_lower):
        return True
    name = rel.name.lower()
    if name in SKIP_FILE_NAMES:
        return True
    if preserve_data and rel.suffix.lower() in SKIP_FILE_SUFFIXES:
        # не затираем локальные xlsx/csv пользователя при обновлении кода из zipball
        return True
    if not include_config and name == "config.yaml":
        return True
    return False


def _iter_code_files(root: Path, include_config: bool) -> Iterable[Path]:
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(root)
        if _should_skip_path(rel, include_config=include_config, preserve_data=True):
            continue
        if rel.suffix.lower() in {".py", ".yaml", ".yml", ".txt", ".command", ".sh", ".md"}:
            yield path
        elif rel.name.upper() == "VERSION":
            yield path
        elif str(rel).replace("\\", "/").startswith(("analyzers/", "tests/")):
            yield path


def _timestamp() -> str:
    from datetime import datetime

    return datetime.now().strftime("%Y%m%d_%H%M%S")


def apply_update_from_zip(
    app_dir: Path,
    zip_url: str,
    *,
    token: Optional[str] = None,
    include_config: bool = False,
    backup: bool = True,
    sha256_url: Optional[str] = None,
    require_sha256: bool = False,
    mode: str = "auto",
) -> Dict[str, Any]:
    """
    mode:
      - release-asset: распаковать содержимое ZIP поверх папки приложения (Windows onedir)
      - release-zipball / branch / auto: копировать только код
    """
    app_dir = Path(app_dir).resolve()
    report: Dict[str, Any] = {
        "copied": [],
        "skipped_config": not include_config,
        "backup": None,
        "sha256_ok": None,
    }

    with tempfile.TemporaryDirectory(prefix="analiz_upd_") as tmp:
        tmp_path = Path(tmp)
        zip_path = tmp_path / "update.zip"
        _http_download(zip_url, zip_path, token=token)

        if sha256_url:
            sha_path = tmp_path / "update.sha256"
            _http_download(sha256_url, sha_path, token=token)
            expected = parse_sha256_text(sha_path.read_text(encoding="utf-8", errors="ignore"))
            if not expected:
                raise RuntimeError("Файл SHA-256 повреждён или неизвестного формата")
            actual = compute_sha256(zip_path)
            if actual != expected:
                raise RuntimeError(
                    f"SHA-256 не совпала (ожидали {expected}, получили {actual})"
                )
            report["sha256_ok"] = actual
        elif require_sha256:
            raise RuntimeError("В релизе нет SHA-256 — обновление отменено")

        extract_dir = tmp_path / "src"
        extract_dir.mkdir()
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)

        children = [p for p in extract_dir.iterdir()]
        # GitHub zipball: одна папка repo-sha/
        if len(children) == 1 and children[0].is_dir() and not (children[0] / EXE_NAME).exists():
            src_root = children[0]
            use_asset_layout = False
        else:
            src_root = extract_dir
            use_asset_layout = (src_root / EXE_NAME).exists() or any(
                p.name.endswith(".exe") for p in src_root.iterdir() if p.is_file()
            )

        if mode == "release-asset" or (mode == "auto" and use_asset_layout):
            # Windows onedir ZIP: всё поверх app_dir, кроме пользовательских данных
            if backup:
                bak_root = app_dir / f".update_backup_{_timestamp()}"
                bak_root.mkdir(parents=True, exist_ok=True)
                report["backup"] = str(bak_root)
            for src in src_root.rglob("*"):
                if not src.is_file():
                    continue
                rel = src.relative_to(src_root)
                if _should_skip_path(rel, include_config=include_config, preserve_data=False):
                    # для asset всё же не трогаем ui_settings / логи
                    if rel.name.lower() in SKIP_FILE_NAMES:
                        continue
                    if not include_config and rel.name.lower() == "config.yaml":
                        # если у пользователя уже есть config — сохраняем
                        if (app_dir / rel).exists():
                            continue
                dest = app_dir / rel
                if backup and dest.exists():
                    bak_dest = Path(report["backup"]) / rel
                    bak_dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(dest, bak_dest)
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)
                report["copied"].append(str(rel))
        else:
            if backup:
                bak_root = app_dir / f".update_backup_{_timestamp()}"
                bak_root.mkdir(parents=True, exist_ok=True)
                report["backup"] = str(bak_root)
            for src in _iter_code_files(src_root, include_config=include_config):
                rel = src.relative_to(src_root)
                dest = app_dir / rel
                if backup and dest.exists():
                    bak_dest = Path(report["backup"]) / rel
                    bak_dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(dest, bak_dest)
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)
                report["copied"].append(str(rel))

    report["count"] = len(report["copied"])
    report["new_version"] = read_local_version(app_dir)
    return report


def format_update_notes(info: UpdateInfo, limit: int = 1200) -> str:
    body = (info.body or "").strip()
    if len(body) > limit:
        body = body[: limit - 1] + "…"
    lines = [
        f"Текущая версия: {info.local_version}",
        f"Доступна: {info.remote_version} ({info.tag})",
        f"Источник: {info.source}",
    ]
    if info.sha256_url:
        lines.append("Проверка SHA-256: да")
    if info.name:
        lines.append(f"Релиз: {info.name}")
    if body:
        lines.append("")
        lines.append(body)
    return "\n".join(lines)
