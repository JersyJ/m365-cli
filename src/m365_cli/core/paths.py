from __future__ import annotations

import re
from pathlib import PureWindowsPath

from m365_cli.core.errors import InvalidInputError

_WSL_MOUNT = re.compile(r"^/mnt/([a-zA-Z])(?:/(.*))?$")


def to_windows_path(path: str) -> str:
    r"""Translate a WSL path so WSL callers can pass /mnt/c/... paths to the Windows-side tool.

    '/mnt/c/Users/x/a.pdf' -> 'C:\Users\x\a.pdf'; Windows paths pass through unchanged.

    Raises:
        InvalidInputError: For Linux-only paths that Windows cannot open.
    """
    p = path.strip().strip('"')
    if m := _WSL_MOUNT.match(p):
        drive, rest = m.group(1).upper(), (m.group(2) or "")
        return str(PureWindowsPath(f"{drive}:/", *rest.split("/")))
    if p.startswith("/"):
        raise InvalidInputError(f"Linux-only path {p!r}: use a path under /mnt/c/... or a Windows path.")
    return str(PureWindowsPath(p))
