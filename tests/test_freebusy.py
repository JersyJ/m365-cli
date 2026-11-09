from datetime import datetime, timedelta

from m365_cli.core.freebusy import Block, BusyStatus, free_windows, parse_freebusy, status_at

DAY = datetime(2026, 11, 10)


def test_parse_freebusy_merges_runs() -> None:
    busy_9_to_10_then_ooo_1030_to_1230 = "0" * 18 + "22" + "0" + "3333"
    blocks = parse_freebusy(busy_9_to_10_then_ooo_1030_to_1230, DAY, 30)
    assert blocks == [
        Block(DAY.replace(hour=9), DAY.replace(hour=10), BusyStatus.BUSY),
        Block(DAY.replace(hour=10, minute=30), DAY.replace(hour=12, minute=30), BusyStatus.OUT_OF_OFFICE),
    ]
    assert status_at(blocks, DAY.replace(hour=11)) == BusyStatus.OUT_OF_OFFICE
    assert status_at(blocks, DAY.replace(hour=13)) == BusyStatus.FREE


def test_free_windows_skip_blocking_but_not_working_elsewhere() -> None:
    busy = [
        Block(DAY.replace(hour=9), DAY.replace(hour=10), BusyStatus.BUSY),
        Block(DAY.replace(hour=10), DAY.replace(hour=11), BusyStatus.WORKING_ELSEWHERE),
    ]
    slots = free_windows(
        busy, DAY.replace(hour=8), DAY.replace(hour=11), timedelta(hours=1), timedelta(minutes=30)
    )
    assert [s.hour + s.minute / 60 for s, _ in slots] == [8.0, 10.0]
