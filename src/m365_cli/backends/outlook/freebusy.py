from __future__ import annotations

from datetime import datetime

import pywintypes

from m365_cli.backends.outlook.session import Com, Outlook, iterate
from m365_cli.core.freebusy import Block, BusyStatus, parse_freebusy
from m365_cli.core.timeutil import from_com, to_com


def recipient_blocks(recipient: Com, day: datetime, minutes: int, days: int) -> list[Block] | None:
    """Free/busy blocks for a resolved Exchange recipient from `day` (midnight) for `days` days.

    `Recipient.FreeBusy(date, minutes, True)` returns one digit per slot from midnight; see
    `core.freebusy`. Returns None when free/busy is not available (e.g. an external address).
    """
    day_start = day.replace(hour=0, minute=0, second=0, microsecond=0)
    try:
        digits = str(recipient.FreeBusy(to_com(day_start), minutes, True))
    except pywintypes.com_error:
        return None
    return parse_freebusy(digits[: days * 24 * 60 // minutes], day_start, minutes)


def own_calendar_blocks(ol: Outlook, start: datetime, end: datetime) -> list[Block]:
    """Blocks from the user's own calendar (more reliable than FreeBusy() on one's own mailbox)."""
    blocks: list[Block] = []
    for item in iterate(ol.calendar_between(start, end)):
        status = BusyStatus(item.BusyStatus)
        s, e = from_com(item.Start), from_com(item.End)
        if status != BusyStatus.FREE and s and e:
            blocks.append(Block(s, e, status))
    return blocks
