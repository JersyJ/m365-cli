from __future__ import annotations

import os
import tomllib
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from m365_cli.core.errors import InvalidInputError

APP_NAME = "m365-cli"

Backend = Literal["outlook", "fake"]


def _app_dir(env_var: str, fallback: str) -> Path:
    base = os.environ.get(env_var)
    return Path(base) / APP_NAME if base else Path.home() / fallback / APP_NAME


def config_dir() -> Path:
    """Directory of config.toml."""
    return _app_dir("APPDATA", "AppData/Roaming")


def data_dir() -> Path:
    """Directory for the ref store, the send log and the fake state."""
    return _app_dir("LOCALAPPDATA", "AppData/Local")


class Settings(BaseModel):
    """User settings (all keys optional).

    Sources, later ones win: the defaults below, `%APPDATA%/m365-cli/config.toml`, then the
    environment variables M365_AGENT_BACKEND, M365_AGENT_FAKE_STATE and M365_AGENT_DATA_DIR.

    `signature_new` / `signature_reply` are "auto" (Outlook's default signature), "none", or a name.

    Example config.toml:

        disabled_features = ["people"]
        list_limit = 20
        body_max_chars = 4000
        signoff = "Sent with my assistant"
        signature_new = "auto"
        signature_reply = "Short"
        work_hours = [8, 17]
        max_sends_per_minute = 5
    """

    model_config = ConfigDict(extra="forbid", frozen=True)

    backend: Backend = "outlook"
    fake_state: Path | None = None
    disabled_features: list[str] = Field(default_factory=list)
    list_limit: int = Field(default=15, ge=1, le=200)
    body_max_chars: int = Field(default=6000, ge=200)
    signoff: str = ""
    signature_new: str = "auto"
    signature_reply: str = "auto"
    work_hours: tuple[int, int] = (8, 17)
    max_sends_per_minute: int = Field(default=5, ge=1, le=60)
    data_dir: Path = Field(default_factory=data_dir)

    @classmethod
    def load(cls, path: Path | None = None, **overrides: Any) -> Settings:
        """Load config.toml and the environment, then apply `overrides` (None values are ignored).

        A `fake_state` without an explicit backend switches to the fake backend.

        Raises:
            InvalidInputError: The file or a value is invalid.
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
            return cls.model_validate(values)
        except (tomllib.TOMLDecodeError, ValidationError) as exc:
            raise InvalidInputError(f"Invalid config {path}: {exc}") from exc


def _environment() -> dict[str, Any]:
    names = {
        "M365_AGENT_BACKEND": "backend",
        "M365_AGENT_FAKE_STATE": "fake_state",
        "M365_AGENT_DATA_DIR": "data_dir",
    }
    return {key: os.environ[var] for var, key in names.items() if os.environ.get(var)}
