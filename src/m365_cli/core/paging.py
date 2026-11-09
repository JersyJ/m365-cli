from __future__ import annotations

import itertools
from collections.abc import Iterable

from pydantic import BaseModel, ConfigDict


class Page[T](BaseModel):
    """One page of results; `has_more` tells the caller to ask for the next page."""

    model_config = ConfigDict(frozen=True)

    items: list[T]
    offset: int = 0
    has_more: bool = False

    @property
    def next_offset(self) -> int:
        """Offset of the next page."""
        return self.offset + len(self.items)


def take_page[T](source: Iterable[T], *, offset: int, limit: int) -> tuple[list[T], bool]:
    """Skip `offset` items and take `limit`. Returns (items, more_follow)."""
    window = list(itertools.islice(source, offset, offset + limit + 1))
    return window[:limit], len(window) > limit


def more_hint[T](page: Page[T]) -> str:
    """Line telling the model how to get the next page ('' when there is none)."""
    if not page.has_more:
        return ""
    return f"[More results: call again with offset={page.next_offset}, or narrow the filter.]"
