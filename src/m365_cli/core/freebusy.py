from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime, timedelta
from enum import IntEnum


class BusyStatus(IntEnum):
    """Free/busy states, with the same values as Outlook.

    A free/busy string has one digit per time slot from midnight: 0 free, 1 tentative, 2 busy,
    3 out of office, 4 working elsewhere. The helpers in this module are pure (no backend).
    """

    FREE = 0
    TENTATIVE = 1
    BUSY = 2
    OUT_OF_OFFICE = 3
    WORKING_ELSEWHERE = 4


STATUS_NAMES = {
    BusyStatus.FREE: "free",
    BusyStatus.TENTATIVE: "tentative",
    BusyStatus.BUSY: "busy",
    BusyStatus.OUT_OF_OFFICE: "out of office",
    BusyStatus.WORKING_ELSEWHERE: "working elsewhere",
}
_BLOCKING = {BusyStatus.BUSY, BusyStatus.OUT_OF_OFFICE, BusyStatus.TENTATIVE}


@dataclass(frozen=True, slots=True)
class Block:
    """One period with one free/busy state."""

    start: datetime
    end: datetime
    status: BusyStatus


def parse_freebusy(digits: str, day_start: datetime, minutes: int) -> list[Block]:
    """Merge runs of equal, non-free digits into blocks."""
    blocks: list[Block] = []
    run_start, run_status = 0, None
    for i, ch in enumerate([*digits, "0"]):
        status = BusyStatus(int(ch)) if ch.isdigit() and int(ch) <= 4 else BusyStatus.FREE
        if status != run_status:
            if run_status not in (None, BusyStatus.FREE):
                start = day_start + timedelta(minutes=run_start * minutes)
                blocks.append(Block(start, day_start + timedelta(minutes=i * minutes), run_status))
            run_start, run_status = i, status
    return blocks


def status_at(blocks: Iterable[Block], at: datetime) -> BusyStatus:
    """State at time `at` (free when no block covers it)."""
    for b in blocks:
        if b.start <= at < b.end:
            return b.status
    return BusyStatus.FREE


def free_windows(
    busy: Iterable[Block], start: datetime, end: datetime, duration: timedelta, step: timedelta
) -> list[tuple[datetime, datetime]]:
    """Slots of `duration` between start and end, stepping by `step`, overlapping no blocking block.

    "Working elsewhere" does not block.
    """
    blocking = sorted((b for b in busy if b.status in _BLOCKING), key=lambda b: b.start)
    slots: list[tuple[datetime, datetime]] = []
    t = start
    while t + duration <= end:
        slot_end = t + duration
        if not any(b.start < slot_end and t < b.end for b in blocking):
            slots.append((t, slot_end))
        t += step
    return slots
