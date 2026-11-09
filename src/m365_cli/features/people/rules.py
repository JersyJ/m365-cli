from __future__ import annotations

from datetime import datetime

from m365_cli.core.freebusy import STATUS_NAMES, Block, BusyStatus, status_at
from m365_cli.features.people.models import Availability


def summarize(blocks: list[Block] | None, at: datetime) -> Availability:
    """Derive the current status, its end, OOO ranges and today's busy blocks (pure function).

    Adjacent OOO blocks are merged (free/busy splits them at day boundaries). None means "unknown".
    """
    if blocks is None:
        return Availability(now="unknown")
    current = status_at(blocks, at)
    until = next((b.end for b in blocks if b.start <= at < b.end), None)
    ooo: list[tuple[datetime, datetime]] = []
    for b in (b for b in blocks if b.status == BusyStatus.OUT_OF_OFFICE and b.end > at):
        if ooo and b.start <= ooo[-1][1]:
            ooo[-1] = (ooo[-1][0], max(ooo[-1][1], b.end))
        else:
            ooo.append((b.start, b.end))
    today = [
        (b.start, b.end)
        for b in blocks
        if b.start.date() == at.date() and b.end > at and b.status in (BusyStatus.BUSY, BusyStatus.TENTATIVE)
    ]
    return Availability(now=STATUS_NAMES[current], current_until=until, out_of_office=ooo, busy_today=today)
