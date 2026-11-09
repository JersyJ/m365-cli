from __future__ import annotations

import json
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from m365_cli.backends.fake.fixtures import build_state
from m365_cli.backends.fake.state import FakeState


class FakeStore:
    """Fake state in memory, written to a JSON file after every change.

    A missing file is created from the synthetic data set.
    """

    def __init__(self, path: Path) -> None:
        """Load the state file at `path`, or create it."""
        self.path = path
        self._lock = threading.RLock()
        if path.exists():
            self._state = FakeState.model_validate_json(path.read_text(encoding="utf-8"))
        else:
            self._state = build_state()
            self._write()

    @property
    def state(self) -> FakeState:
        """Current state (read-only by convention; change it through `edit()`)."""
        return self._state

    @contextmanager
    def edit(self) -> Iterator[FakeState]:
        """Context manager for changes; saves on success, discards everything if the block raises."""
        with self._lock:
            draft = self._state.model_copy(deep=True)
            yield draft
            self._state = draft
            self._write()

    def _write(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._state.model_dump(mode="json"), ensure_ascii=False, indent=1), "utf-8")
        tmp.replace(self.path)


def write_state(path: Path, state: FakeState) -> None:
    """Write a complete state to `path` (e.g. a snapshot)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state.model_dump(mode="json"), ensure_ascii=False, indent=1), "utf-8")
