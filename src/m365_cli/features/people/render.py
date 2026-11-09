from __future__ import annotations

from datetime import datetime

from m365_cli.core.timeutil import fmt
from m365_cli.features.people.models import Availability, Person


def _ranges(ranges: list[tuple[datetime, datetime]]) -> str:
    return ", ".join(f"{fmt(s)}-{fmt(e, base=s)}" for s, e in ranges)


def person(p: Person, availability: Availability | None = None) -> str:
    """Contact card, plus availability when given."""
    head = p.name + (f" <{p.email}>" if p.email else "")
    details = " | ".join(x for x in (p.title, p.department, p.office) if x)
    lines = [head] + ([details] if details else [])
    contact = " | ".join(x for x in (p.phone and f"tel {p.phone}", p.mobile and f"mob {p.mobile}") if x)
    if contact:
        lines.append(contact)
    if p.manager:
        lines.append(f"Manager: {p.manager}")
    if availability is not None:
        status = availability.now
        if availability.current_until and status != "free":
            status += f" until {fmt(availability.current_until)}"
        lines.append(f"Now: {status}")
        if availability.out_of_office:
            lines.append("Out of office: " + _ranges(availability.out_of_office))
        if availability.busy_today:
            lines.append("Busy today: " + _ranges(availability.busy_today))
    return "\n".join(lines)
