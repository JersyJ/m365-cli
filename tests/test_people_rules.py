from datetime import datetime

from m365_cli.core.freebusy import Block, BusyStatus
from m365_cli.features.people.rules import summarize

NOW = datetime(2026, 11, 10, 9, 15)


def test_summarize_reports_current_status_and_merged_ooo() -> None:
    blocks = [
        Block(datetime(2026, 11, 10, 9), datetime(2026, 11, 10, 10), BusyStatus.BUSY),
        Block(datetime(2026, 11, 11), datetime(2026, 11, 12), BusyStatus.OUT_OF_OFFICE),
        Block(datetime(2026, 11, 12), datetime(2026, 11, 13), BusyStatus.OUT_OF_OFFICE),
    ]
    a = summarize(blocks, NOW)
    assert a.now == "busy" and a.current_until == datetime(2026, 11, 10, 10)
    assert a.out_of_office == [(datetime(2026, 11, 11), datetime(2026, 11, 13))]
    assert a.busy_today == [(datetime(2026, 11, 10, 9), datetime(2026, 11, 10, 10))]


def test_summarize_unknown_without_freebusy() -> None:
    assert summarize(None, NOW).now == "unknown"
