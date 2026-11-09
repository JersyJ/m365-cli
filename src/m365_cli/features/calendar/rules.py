from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime, timedelta

from m365_cli.core.errors import InvalidInputError
from m365_cli.core.freebusy import Block, free_windows
from m365_cli.features.calendar.models import Repeat, Slot

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


WEEKDAY_BITS = {"sun": 1, "mon": 2, "tue": 4, "wed": 8, "thu": 16, "fri": 32, "sat": 64}
"""Outlook's `OlDaysOfWeek` bit for each weekday."""


def check_repeat(repeat: Repeat, start: datetime) -> list[str]:
    """Validate a repeat rule and return its weekdays (empty unless every='week').

    Raises:
        InvalidInputError: Both `until` and `count`, an `until` before the start, or weekdays for days/months.
    """
    if repeat.until is not None and repeat.count is not None:
        raise InvalidInputError("Give either until or count, not both.")
    if repeat.until is not None and repeat.until < start.date():
        raise InvalidInputError("until is before the start.")
    if repeat.every != "week":
        if repeat.days:
            raise InvalidInputError("days only works with every='week'.")
        return []
    return list(repeat.days or [list(WEEKDAY_BITS)[(start.weekday() + 1) % 7]])


def describe_repeat(repeat: Repeat, start: datetime) -> str:
    """Short text such as 'every 2 weeks on Mon, Wed until 2026-12-31'."""
    unit = repeat.every if repeat.interval == 1 else f"{repeat.interval} {repeat.every}s"
    days = check_repeat(repeat, start)
    on = f" on {', '.join(d.capitalize() for d in days)}" if days else ""
    end = f" until {repeat.until}" if repeat.until else (f", {repeat.count} times" if repeat.count else "")
    return f"every {unit}{on}{end}"
