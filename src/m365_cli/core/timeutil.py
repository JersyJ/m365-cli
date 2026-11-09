from __future__ import annotations

import re
from datetime import UTC, date, datetime, time, timedelta

from m365_cli.core.errors import InvalidInputError

_REL = re.compile(r"^([+-]?)(\d+)\s*([mhdw])$")
_UNITS = {"m": "minutes", "h": "hours", "d": "days", "w": "weeks"}
_WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
_CLOCK = re.compile(r"^(\d{1,2})(?::(\d{2}))?$")


class _Clock:
    """Holds the fixed clock of the process; None means real time."""

    fixed: datetime | None = None


WHEN_FORMATS = (
    "'2026-11-12 09:30', '2026-11-12', 'today', 'tomorrow 14:00', 'fri 9:30', '14:00', '+2h', 'now'"
)


def now() -> datetime:
    """Current local time without microseconds (or the fixed clock, see `use_clock`).

    All datetimes inside the app are naive local time, which is what Outlook's object model uses.
    """
    return _Clock.fixed or datetime.now().replace(microsecond=0)


def use_clock(value: datetime | None) -> None:
    """Freeze `now()` for this process; None restores the real clock.

    The fake backend uses this so that "today" means the day of its data.
    """
    _Clock.fixed = value


def from_com(value: datetime | None) -> datetime | None:
    """pywin32 labels Outlook times as UTC although they are local wall-clock times: drop the label.

    Outlook uses 4501-01-01 as "no date"; that becomes None.
    """
    if value is None:
        return None
    naive = value.replace(tzinfo=None)
    return None if naive.year >= 4500 else naive


def to_com(value: datetime) -> datetime:
    """Prepare a naive local datetime for COM.

    pywin32 converts naive datetimes from local time to UTC before passing them on, but Outlook
    expects local wall-clock time. Labelling the value as UTC makes pywin32 pass it through unchanged.
    """
    return value.replace(tzinfo=UTC)


def to_dasl_utc(value: datetime) -> str:
    """Format a naive local datetime for a DASL (@SQL=) filter, which compares in UTC."""
    return value.astimezone(UTC).strftime("%Y-%m-%d %H:%M")


def _parse_day(token: str, today: date) -> date | None:
    t = token.lower()
    if t == "today":
        return today
    if t == "tomorrow":
        return today + timedelta(days=1)
    if t == "yesterday":
        return today - timedelta(days=1)
    if t[:3] in _WEEKDAYS and t.isalpha():
        delta = (_WEEKDAYS.index(t[:3]) - today.weekday()) % 7 or 7
        return today + timedelta(days=delta)
    try:
        return date.fromisoformat(token)
    except ValueError:
        return None


def _parse_clock(text: str) -> time:
    m = _CLOCK.match(text)
    try:
        if m is None:
            raise ValueError(text)
        return time(int(m.group(1)), int(m.group(2) or 0))
    except ValueError as exc:
        raise InvalidInputError(f"Unrecognized time: {text!r} (try '9:30' or '14:00').") from exc


def parse_when(text: str, *, base: datetime | None = None, default_time: time = time(0, 0)) -> datetime:
    """Parse '2026-11-10 14:00', '2026-11-10', 'tomorrow 9:30', 'fri 14:00', '14:00', '+2h', '-3d', 'now'.

    A weekday means the next such day, never today. `base` is the reference for relative values
    (default: now); `default_time` is used when only a date is given.

    Raises:
        InvalidInputError: Anything else.
    """
    base = base or now()
    s = text.strip()
    if not s:
        raise InvalidInputError(f"Empty date/time. Try e.g. {WHEN_FORMATS}.")
    if s.lower() == "now":
        return base
    if m := _REL.match(s):
        sign = -1 if m.group(1) == "-" else 1
        return base + sign * timedelta(**{_UNITS[m.group(3)]: int(m.group(2))})
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        pass
    parts = s.replace("T", " ").split()
    day = _parse_day(parts[0], base.date())
    clock = parts[1] if day is not None and len(parts) == 2 else (parts[0] if day is None else None)
    if day is None and (len(parts) != 1 or not _CLOCK.match(parts[0])):
        raise InvalidInputError(f"Unrecognized date/time: {text!r}. Try e.g. {WHEN_FORMATS}.")
    t = _parse_clock(clock) if clock is not None else default_time
    return datetime.combine(day or base.date(), t)


def parse_since(text: str, *, base: datetime | None = None) -> datetime:
    """'2d', '12h', '1w' mean that long ago; anything else is parsed by `parse_when`."""
    base = base or now()
    if (m := _REL.match(text.strip())) and m.group(1) == "":
        return base - timedelta(**{_UNITS[m.group(3)]: int(m.group(2))})
    return parse_when(text, base=base)


def fmt(value: datetime | None, *, base: datetime | None = None) -> str:
    """Compact display: '14:05' today, 'Mon 11-10 14:05' within a week, else '2026-11-10 14:05'."""
    if value is None:
        return "-"
    base = base or now()
    if value.date() == base.date():
        return value.strftime("%H:%M")
    if abs((value.date() - base.date()).days) < 7:
        return value.strftime("%a %m-%d %H:%M")
    return value.strftime("%Y-%m-%d %H:%M")
