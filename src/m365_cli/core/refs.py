from __future__ import annotations

import base64
import hashlib
import sqlite3
import threading
from dataclasses import dataclass
from pathlib import Path

from m365_cli.core.errors import InvalidInputError, NotFoundError

_HASH_CHARS = 7


@dataclass(frozen=True, slots=True)
class BackendId:
    """Identifies an item in a backend.

    Attributes:
        kind: "m" mail item, "e" calendar item.
        item_id: e.g. the Outlook EntryID.
        container_id: e.g. the Outlook StoreID.
    """

    kind: str
    item_id: str
    container_id: str = ""


class RefStore:
    """Short, stable references for long backend IDs.

    Outlook EntryIDs are ~140 hex chars; repeating them costs tokens and invites copy errors. We map
    each one to a short ref such as `m:7KQ2PXA` (kind prefix + 7 base32 chars of a SHA-1). Refs are
    deterministic, so the same item always gets the same ref, and they are persisted in a small
    SQLite file so they survive restarts.
    """

    def __init__(self, path: Path) -> None:
        """Open or create the store at `path`."""
        path.parent.mkdir(parents=True, exist_ok=True)
        self._path = path
        self._lock = threading.Lock()
        with self._connect() as db:
            db.execute(
                "CREATE TABLE IF NOT EXISTS refs (ref TEXT PRIMARY KEY, kind TEXT NOT NULL, "
                "item_id TEXT NOT NULL, container_id TEXT NOT NULL, "
                "touched TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)"
            )

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self._path, timeout=5)

    @staticmethod
    def make_ref(kind: str, item_id: str) -> str:
        """Deterministic ref for an item ID."""
        digest = hashlib.sha1(item_id.encode(), usedforsecurity=False).digest()
        return f"{kind}:{base64.b32encode(digest).decode()[:_HASH_CHARS]}"

    def put(self, backend_id: BackendId) -> str:
        """Store the mapping and return the ref."""
        ref = self.make_ref(backend_id.kind, backend_id.item_id)
        with self._lock, self._connect() as db:
            db.execute(
                "INSERT INTO refs (ref, kind, item_id, container_id) VALUES (?, ?, ?, ?) "
                "ON CONFLICT(ref) DO UPDATE SET item_id = excluded.item_id, "
                "container_id = excluded.container_id, touched = CURRENT_TIMESTAMP",
                (ref, backend_id.kind, backend_id.item_id, backend_id.container_id),
            )
        return ref

    def get(self, ref: str, *, kinds: str | None = None) -> BackendId:
        """Look up a ref; `kinds` restricts the allowed prefixes (e.g. "em").

        Raises:
            InvalidInputError: Malformed ref or wrong kind.
            NotFoundError: Unknown ref.
        """
        ref = ref.strip()
        if ":" not in ref:
            raise InvalidInputError(f"Not a ref: {ref!r} (expected e.g. 'm:7KQ2PXA').")
        kind = ref.split(":", 1)[0]
        if kinds is not None and kind not in kinds:
            raise InvalidInputError(f"Ref {ref} has kind '{kind}', expected one of: {', '.join(kinds)}.")
        with self._lock, self._connect() as db:
            row = db.execute("SELECT kind, item_id, container_id FROM refs WHERE ref = ?", (ref,)).fetchone()
        if row is None:
            raise NotFoundError(f"Unknown ref {ref}; list items again to get current refs.")
        return BackendId(*row)

    def prune(self, keep_days: int = 60) -> int:
        """Delete refs not used for `keep_days`; returns how many were deleted."""
        with self._lock, self._connect() as db:
            cur = db.execute("DELETE FROM refs WHERE touched < datetime('now', ?)", (f"-{keep_days} days",))
            return cur.rowcount
