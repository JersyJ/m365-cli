from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime, timedelta

from m365_cli.core.errors import InvalidInputError
from m365_cli.core.freebusy import Block, free_windows
from m365_cli.features.calendar.models import Slot

SLOT_STEP = timedelta(minutes=30)
MAX_DAYS = 14


def check_period(start: datetime, end: datetime) -> None:
    """Raise InvalidInputError unless `end` is after `start`."""
    if end <= start:
        raise InvalidInputError("End must be after start.")


def day_start(value: datetime) -> datetime:
    """Midnight of the day of `value`."""
    return value.replace(hour=0, minute=0, second=0, microsecond=0)


def round_up(value: datetime, step: timedelta) -> datetime:
    """Round `value` up to the next multiple of `step` after midnight."""
    base = day_start(value)
    steps = -(-(value - base) // step)
    return base + steps * step


def free_slots(
    blocks: Iterable[Block],
    *,
    first_day: datetime,
    days: int,
    duration_minutes: int,
    work_hours: tuple[int, int],
    earliest: datetime,
    limit: int,
) -> list[Slot]:
    """Free slots within working hours on weekdays, shared by all backends.

    Args:
        blocks: Busy blocks of everyone involved.
        first_day: First day to look at.
        days: Number of days (1-14).
        duration_minutes: Length of each slot.
        work_hours: (start hour, end hour) of a working day.
        earliest: No slot starts before this (usually now).
        limit: Maximum number of slots.

    Raises:
        InvalidInputError: `days` is out of range.
    """
    if not 1 <= days <= MAX_DAYS:
        raise InvalidInputError(f"days must be between 1 and {MAX_DAYS}.")
    blocks = list(blocks)
    start_hour, end_hour = work_hours
    duration = timedelta(minutes=duration_minutes)
    earliest = round_up(earliest, SLOT_STEP)
    slots: list[Slot] = []
    for offset in range(days):
        date = day_start(first_day) + timedelta(days=offset)
        if date.weekday() >= 5:
            continue
        window_start = max(date.replace(hour=start_hour), earliest)
        for s, e in free_windows(blocks, window_start, date.replace(hour=end_hour), duration, SLOT_STEP):
            slots.append(Slot(start=s, end=e))
            if len(slots) >= limit:
                return slots
    return slots
