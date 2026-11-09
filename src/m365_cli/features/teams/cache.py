from __future__ import annotations

import json
import threading
from pathlib import Path

from m365_cli.features.teams.models import Member


class ProfileCache:
    """Display names and emails of Teams users by user ID (MRI), kept in a local JSON file.

    1:1 chats have no title, so their title is the other person's name; the cache avoids one profile
    request per chat on every listing.
    """

    def __init__(self, path: Path) -> None:
        """Use the file at `path` (created on the first save)."""
        self._path = path
        self._lock = threading.Lock()
        self._data: dict[str, dict[str, str]] | None = None

    def _load(self) -> dict[str, dict[str, str]]:
        if self._data is None:
            data: dict[str, dict[str, str]] = {}
            try:
                if self._path.exists():
                    data = json.loads(self._path.read_text(encoding="utf-8"))
            except ValueError:
                data = {}
            self._data = data
        return self._data

    def get(self, mri: str) -> Member | None:
        """The cached person, or None."""
        with self._lock:
            entry = self._load().get(mri)
        return Member(name=entry["name"], mri=mri, email=entry.get("email", "")) if entry else None

    def put_all(self, members: list[Member]) -> None:
        """Add or update people and save the file."""
        if not members:
            return
        with self._lock:
            data = self._load()
            for m in members:
                data[m.mri] = {"name": m.name, "email": m.email}
            self._path.parent.mkdir(parents=True, exist_ok=True)
            self._path.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")
