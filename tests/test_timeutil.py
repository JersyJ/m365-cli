from datetime import UTC, datetime, time, timedelta, timezone

import pytest

from m365_cli.core.errors import InvalidInputError
from m365_cli.core.timeutil import fmt, from_com, parse_since, parse_when, to_dasl_utc

MONDAY_AFTERNOON = datetime(2026, 11, 9, 15, 30)
BASE = MONDAY_AFTERNOON


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("2026-11-12 09:30", datetime(2026, 11, 12, 9, 30)),
        ("2026-11-12", datetime(2026, 11, 12)),
        ("today", datetime(2026, 11, 9)),
        ("tomorrow 14:00", datetime(2026, 11, 10, 14, 0)),
        ("tomorrow 9", datetime(2026, 11, 10, 9, 0)),
        ("fri 9:30", datetime(2026, 11, 13, 9, 30)),
        ("mon", datetime(2026, 11, 16)),
        ("14:00", datetime(2026, 11, 9, 14, 0)),
        ("+2h", BASE + timedelta(hours=2)),
        ("-3d", BASE - timedelta(days=3)),
        ("now", BASE),
    ],
)
def test_parse_when(text: str, expected: datetime) -> None:
    assert parse_when(text, base=BASE) == expected


def test_same_weekday_means_next_week() -> None:
    assert parse_when("mon", base=MONDAY_AFTERNOON) == datetime(2026, 11, 16)


def test_parse_when_default_time() -> None:
    assert parse_when("tomorrow", base=BASE, default_time=time(8)) == datetime(2026, 11, 10, 8)


@pytest.mark.parametrize("bad", ["", "someday", "tomorrow 25:00", "a b c"])
def test_parse_when_rejects_garbage(bad: str) -> None:
    with pytest.raises(InvalidInputError):
        parse_when(bad, base=BASE)


def test_parse_since_relative_means_ago() -> None:
    assert parse_since("2d", base=BASE) == BASE - timedelta(days=2)
    assert parse_since("yesterday", base=BASE) == datetime(2026, 11, 8)


def test_from_com_drops_misleading_utc_label_and_null_date() -> None:
    labelled = datetime(2026, 11, 9, 15, 18, tzinfo=UTC)
    assert from_com(labelled) == datetime(2026, 11, 9, 15, 18)
    assert from_com(datetime(4501, 1, 1)) is None
    assert from_com(None) is None


def test_to_dasl_utc_converts_local_to_utc() -> None:
    local = datetime(2026, 11, 9, 12, 0, tzinfo=timezone(timedelta(hours=1)))
    assert to_dasl_utc(local) == "2026-11-09 11:00"


def test_fmt_is_compact_relative_to_base() -> None:
    assert fmt(datetime(2026, 11, 9, 8, 5), base=BASE) == "08:05"
    assert fmt(datetime(2026, 11, 10, 8, 5), base=BASE) == "Tue 11-10 08:05"
    assert fmt(datetime(2026, 1, 2, 8, 5), base=BASE) == "2026-01-02 08:05"
    assert fmt(None) == "-"


def test_to_com_keeps_wall_clock_and_labels_utc() -> None:
    from m365_cli.core.timeutil import to_com

    value = to_com(datetime(2026, 11, 14, 10, 0))
    assert value.tzinfo is UTC and (value.hour, value.minute) == (10, 0)
