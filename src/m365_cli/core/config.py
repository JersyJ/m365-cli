from __future__ import annotations

import os
import tomllib
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from m365_cli.core.errors import InvalidInputError

APP_NAME = "m365-cli"

Backend = Literal["outlook", "web", "fake"]
Window = Literal["offscreen", "minimized", "normal"]


def _app_dir(env_var: str, fallback: str) -> Path:
    base = os.environ.get(env_var)
    return Path(base) / APP_NAME if base else Path.home() / fallback / APP_NAME


def config_dir() -> Path:
    """Directory of config.toml: %APPDATA%/m365-cli on Windows, else $XDG_CONFIG_HOME (~/.config)/m365-cli."""
    if os.name == "nt":
        return _app_dir("APPDATA", "AppData/Roaming")
    return _app_dir("XDG_CONFIG_HOME", ".config")


def data_dir() -> Path:
    """Directory for refs, the send log, the browser profile and the fake state.

    %LOCALAPPDATA%/m365-cli on Windows, else $XDG_DATA_HOME (~/.local/share)/m365-cli.
    """
    if os.name == "nt":
        return _app_dir("LOCALAPPDATA", "AppData/Local")
    return _app_dir("XDG_DATA_HOME", ".local/share")


class Settings(BaseModel):
    """User settings (all keys optional).

    Sources, later ones win: the defaults below, `config.toml` in `config_dir()`, then the
    environment variables M365_AGENT_BACKEND, M365_AGENT_FAKE_STATE and M365_AGENT_DATA_DIR.

    `backend` is "outlook" (classic Outlook through COM, Windows only), "web" (Outlook on the web in the
    dedicated browser) or "fake".

    `signature_new` / `signature_reply` are "auto" (Outlook's default signature), "none", or a name.
    With classic Outlook, `signature_reply` has no effect: on a reply or forward Outlook inserts its own
    reply signature. With the web backend both are your roaming (cloud) signatures.

    `mail_style` is the font of the text the tool writes, as CSS (e.g. "font-family:Aptos;
    font-size:11pt"). Empty: classic Outlook's compose font, or the web compose font when one is set.

    `read_protected` lets the web backend show the decrypted text of rights-protected mail. Off by
    default: such mail is shown as [protected] without its content.

    Example config.toml:

        disabled_features = ["people"]
        list_limit = 20
        body_max_chars = 4000
        signoff = "Sent with my assistant"
        signature_new = "auto"
        signature_reply = "auto"
        work_hours = [8, 17]
        max_sends_per_minute = 5
        teams_protected = ["All Company*", "*Announcements*"]

    Browser (Teams and the web backend): `teams_port` is the DevTools port of the dedicated browser,
    `teams_browser` its executable (default: Chrome, then Edge), `teams_profile_dir` its profile
    (default: data_dir / chrome-teams), `teams_window` where its window starts ("offscreen",
    "minimized" or "normal"), and `teams_protected` lists chat titles ('*' wildcards) the tool never
    sends to.
    """

    model_config = ConfigDict(extra="forbid", frozen=True)

    backend: Backend = "web"
    fake_state: Path | None = None
    disabled_features: list[str] = Field(default_factory=list)
    list_limit: int = Field(default=15, ge=1, le=200)
    body_max_chars: int = Field(default=6000, ge=200)
    signoff: str = ""
    signature_new: str = "auto"
    signature_reply: str = "auto"
    mail_style: str = ""
    read_protected: bool = False
    work_hours: tuple[int, int] = (8, 17)
    max_sends_per_minute: int = Field(default=5, ge=1, le=60)
    data_dir: Path = Field(default_factory=data_dir)
    teams_port: int = Field(default=9222, ge=1024, le=65535)
    teams_browser: Path | None = None
    teams_profile_dir: Path | None = None
    teams_window: Window = "offscreen"
    teams_protected: list[str] = Field(default_factory=list)

    @classmethod
    def load(cls, path: Path | None = None, **overrides: Any) -> Settings:
        """Load config.toml and the environment, then apply `overrides` (None values are ignored).

        A `fake_state` without an explicit backend switches to the fake backend.

        Raises:
            InvalidInputError: The file or a value is invalid, or classic Outlook is chosen off Windows.
        """
        path = path or config_dir() / "config.toml"
        values: dict[str, Any] = {}
        try:
            if path.exists():
                values = tomllib.loads(path.read_text(encoding="utf-8"))
            values |= _environment()
            values |= {k: v for k, v in overrides.items() if v is not None}
            if values.get("fake_state") and "backend" not in overrides:
                values.setdefault("backend", "fake")
            settings = cls.model_validate(values)
        except (tomllib.TOMLDecodeError, ValidationError) as exc:
            raise InvalidInputError(f"Invalid config {path}: {exc}") from exc
        if settings.backend == "outlook" and os.name != "nt":
            raise InvalidInputError("backend = 'outlook' (classic Outlook) works only on Windows; use 'web'.")
        return settings


def _environment() -> dict[str, Any]:
    names = {
        "M365_AGENT_BACKEND": "backend",
        "M365_AGENT_FAKE_STATE": "fake_state",
        "M365_AGENT_DATA_DIR": "data_dir",
    }
    return {key: os.environ[var] for var, key in names.items() if os.environ.get(var)}
