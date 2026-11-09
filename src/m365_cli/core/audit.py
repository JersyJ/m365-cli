from __future__ import annotations

import sqlite3
import threading
from dataclasses import dataclass
from pathlib import Path

from m365_cli.core.errors import RateLimitError

_WINDOW_SECONDS = 60


@dataclass(frozen=True, slots=True)
class AuditEntry:
    """One row of the send log."""

    at: str
    action: str
    summary: str


class AuditLog:
    """Send log with a rate limit.

    Every action that notifies other people (mail, invitation, meeting response, cancellation) calls
    `guard()` before and `record()` after it. The log is a SQLite file, so the limit holds across
    processes (CLI and MCP server).
    """

    def __init__(self, path: Path, *, max_per_minute: int) -> None:
        """Open or create the log at `path`."""
        path.parent.mkdir(parents=True, exist_ok=True)
        self._path = path
        self._max = max_per_minute
        self._lock = threading.Lock()
        with self._connect() as db:
            db.execute(
                "CREATE TABLE IF NOT EXISTS sends (at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')), "
                "action TEXT NOT NULL, summary TEXT NOT NULL)"
            )

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self._path, timeout=5)

    def guard(self, action: str) -> None:
        """Refuse when one more send would exceed the per-minute limit.

        Raises:
            RateLimitError: There were already `max_per_minute` sends in the last 60 seconds.
        """
        with self._lock, self._connect() as db:
            (count,) = db.execute(
                "SELECT COUNT(*) FROM sends WHERE at > datetime('now', 'localtime', ?)",
                (f"-{_WINDOW_SECONDS} seconds",),
            ).fetchone()
        if count >= self._max:
            raise RateLimitError(
                f"Send limit reached ({self._max} per minute). Wait a minute, then retry '{action}'."
            )

    def record(self, action: str, summary: str) -> None:
        """Log one completed send."""
        with self._lock, self._connect() as db:
            db.execute("INSERT INTO sends (action, summary) VALUES (?, ?)", (action, summary))

    def recent(self, limit: int = 20) -> list[AuditEntry]:
        """Last send actions, newest first."""
        with self._lock, self._connect() as db:
            rows = db.execute(
                "SELECT at, action, summary FROM sends ORDER BY rowid DESC LIMIT ?", (limit,)
            ).fetchall()
        return [AuditEntry(*row) for row in rows]
