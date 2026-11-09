from __future__ import annotations

import json
import threading
from datetime import timedelta
from pathlib import Path

from m365_cli.core.errors import NotFoundError
from m365_cli.core.timeutil import now
from m365_cli.features.teams.models import TeamsDraft

_KEEP = timedelta(days=7)


class DraftFile:
    """Teams drafts of the real backend, kept in a local JSON file until they are sent.

    Teams has no draft we could use for this, so the preview step stays on this computer. Drafts
    older than 7 days are dropped.
    """

    def __init__(self, path: Path) -> None:
        """Use the file at `path` (created on the first save)."""
        self._path = path
        self._lock = threading.Lock()

    def _load(self) -> dict[str, TeamsDraft]:
        if not self._path.exists():
            return {}
        raw = json.loads(self._path.read_text(encoding="utf-8") or "{}")
        drafts = {k: TeamsDraft.model_validate(v) for k, v in raw.items()}
        limit = now() - _KEEP
        return {k: d for k, d in drafts.items() if d.created >= limit}

    def _save(self, drafts: dict[str, TeamsDraft]) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._path.with_suffix(".tmp")
        data = {k: d.model_dump(mode="json") for k, d in drafts.items()}
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(self._path)

    def put(self, draft: TeamsDraft) -> None:
        """Save or replace a draft."""
        with self._lock:
            drafts = self._load()
            drafts[draft.id] = draft
            self._save(drafts)

    def get(self, draft_id: str) -> TeamsDraft:
        """The draft with this ID.

        Raises:
            NotFoundError: Unknown, already sent, or older than 7 days.
        """
        with self._lock:
            draft = self._load().get(draft_id)
        if draft is None:
            raise NotFoundError("Unknown Teams draft (already sent, or older than 7 days). Draft it again.")
        return draft

    def remove(self, draft_id: str) -> None:
        """Delete a draft (after it was sent)."""
        with self._lock:
            drafts = self._load()
            if drafts.pop(draft_id, None) is not None:
                self._save(drafts)
