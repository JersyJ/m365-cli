from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from collections.abc import Callable, Iterable, Mapping
from pathlib import Path

from m365_cli import __version__
from m365_cli.core.errors import M365Error

REPO = "JersyJ/m365-cli"
_LATEST_URL = f"https://api.github.com/repos/{REPO}/releases/latest"


def parse_version(text: str) -> tuple[int, ...] | None:
    """Read 'v0.2.0' or '0.2.0' as (0, 2, 0); None when it is not a dotted number."""
    parts = text.strip().removeprefix("v").split(".")
    try:
        return tuple(int(p) for p in parts)
    except ValueError:
        return None


def is_newer(tag: str, installed: str) -> bool:
    """True when the release tag is a higher version than the installed one."""
    tag_v, cur_v = parse_version(tag), parse_version(installed)
    if tag_v is None or cur_v is None:
        return tag.removeprefix("v") != installed
    return tag_v > cur_v


def pick_wheel(assets: Iterable[Mapping[str, object]]) -> Mapping[str, object] | None:
    """The first .whl asset of a release, or None."""
    return next((a for a in assets if str(a.get("name", "")).endswith(".whl")), None)


def source_checkout() -> Path | None:
    """The repo root when this code runs from a git checkout (editable install), else None."""
    root = Path(__file__).resolve().parents[3]
    return root if (root / ".git").is_dir() else None


def _token() -> str:
    for name in ("GH_TOKEN", "GITHUB_TOKEN"):
        if value := os.environ.get(name):
            return value
    if gh := (shutil.which("gh") or shutil.which("gh.exe")):
        result = subprocess.run([gh, "auth", "token"], capture_output=True, text=True, check=False)
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    raise M365Error(
        "No GitHub token: the repo is private. Install GitHub CLI and run `gh auth login`, "
        "or set GH_TOKEN to a token with read access to " + REPO + "."
    )


def _api(url: str, token: str, accept: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": accept,
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise M365Error(
                f"No release found on {REPO} (or the token has no access to the repo)."
            ) from None
        raise M365Error(f"GitHub answered {exc.code} for {url}.") from None
    except urllib.error.URLError as exc:
        raise M365Error(f"Cannot reach GitHub: {exc.reason}") from None


def _find_uv() -> str | None:
    """Find uv.exe, also when started from WSL (then the Windows PATH is not visible)."""
    if found := (shutil.which("uv") or shutil.which("uv.exe")):
        return found
    parts = Path(sys.executable).resolve().parts
    if "AppData" not in parts:
        return None
    home = Path(*parts[: parts.index("AppData")])
    candidates = [
        home / "AppData/Local/Microsoft/WinGet/Links/uv.exe",
        *sorted((home / "AppData/Local/Microsoft/WinGet/Packages").glob("astral-sh.uv_*/uv.exe")),
        home / ".local/bin/uv.exe",
    ]
    return next((str(c) for c in candidates if c.is_file()), None)


def self_update(echo: Callable[[str], None]) -> None:
    """Update a wheel install to the latest GitHub release.

    Checks the latest release, downloads its wheel and reinstalls it with
    `uv tool install --reinstall`. The reinstall runs detached after a short
    delay, because Windows locks the running m365.exe.
    """
    if repo := source_checkout():
        raise M365Error(
            f"This is a source install at {repo}. Update it from WSL with: bash scripts/update.sh"
        )
    uv = _find_uv()
    if not uv:
        raise M365Error("uv not found on PATH. Install it first: winget install astral-sh.uv")

    token = _token()
    release = json.loads(_api(_LATEST_URL, token, "application/vnd.github+json"))
    tag = str(release.get("tag_name", ""))
    if not is_newer(tag, __version__):
        echo(f"Already up to date ({__version__} installed, latest release is {tag}).")
        return
    asset = pick_wheel(release.get("assets", []))
    if asset is None:
        raise M365Error(f"Release {tag} has no .whl asset.")

    echo(f"Downloading {asset['name']} ({tag}) ...")
    tmp = Path(tempfile.gettempdir())
    wheel = tmp / str(asset["name"])
    wheel.write_bytes(_api(str(asset["url"]), token, "application/octet-stream"))

    log = tmp / "m365-update.log"
    if os.name == "nt":
        # The running m365.exe locks its own files, so reinstall detached after this process exits.
        script = tmp / "m365-update.bat"
        script.write_text(
            "@echo off\r\n"
            "timeout /t 3 /nobreak >nul\r\n"
            f'"{uv}" tool install --reinstall "{wheel}" > "{log}" 2>&1\r\n',
            encoding="utf-8",
        )
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NEW_PROCESS_GROUP", 0
        )
        subprocess.Popen(
            ["cmd.exe", "/c", str(script)],
            creationflags=flags,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
        echo(f"Installing {tag} in the background (a few seconds). Log: {log}")
        echo("Check with: m365 version")
    else:
        subprocess.run([uv, "tool", "install", "--reinstall", str(wheel)], check=True)
        echo(f"Installed {tag}.")
