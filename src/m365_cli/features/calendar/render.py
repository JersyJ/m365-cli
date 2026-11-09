from __future__ import annotations

from datetime import datetime

from m365_cli.core.paging import Page, more_hint
from m365_cli.core.text import one_line, untrusted
from m365_cli.core.timeutil import fmt
from m365_cli.features.calendar.models import Event, EventSummary, FreeTimeResult, MeetingDraft, Slot


def _span(start: datetime, end: datetime, all_day: bool = False) -> str:
    if all_day:
        days = (end.date() - start.date()).days
        return start.strftime("%a %m-%d") + (" all day" if days <= 1 else f" +{days}d all day")
    same_day = start.date() == end.date()
    return f"{fmt(start)}-{end:%H:%M}" if same_day else f"{fmt(start)} → {fmt(end)}"


def event_line(e: EventSummary) -> str:
    """One line per event: ref, time, subject, location, show-as/response/flags."""
    parts = [f"{e.ref} {_span(e.start, e.end, e.all_day)}", one_line(e.subject, 70)]
    if e.location:
        parts.append(one_line(e.location, 30))
    status = [e.show_as]
    if e.my_response and e.my_response != "accepted":
        status.append(e.my_response)
    if e.recurring:
        status.append("recurring")
    if e.join_url:
        status.append("teams")
    if e.canceled:
        status.append("CANCELED")
    return " | ".join(parts) + " | " + ", ".join(status)


def events(page: Page[EventSummary]) -> str:
    """All lines of a page, plus the hint for the next page if there is one."""
    if not page.items:
        return "(no events)"
    lines = [event_line(e) for e in page.items]
    hint = more_hint(page)
    return "\n".join([*lines, hint] if hint else lines)


def event(e: Event, *, mark_untrusted: bool = False) -> str:
    """Event details. `mark_untrusted` wraps the description in markers (for the model)."""
    lines = [event_line(e)]
    if e.organizer:
        lines.append(f"Organizer: {e.organizer}")
    if e.required:
        lines.append(f"Required: {e.required}")
    if e.optional:
        lines.append(f"Optional: {e.optional}")
    if e.join_url:
        lines.append(f"Join: {e.join_url}")
    if e.body:
        body = e.body + (" […]" if e.truncated else "")
        lines += ["", untrusted("EVENT DESCRIPTION", body) if mark_untrusted else body]
    return "\n".join(lines)


def draft(d: MeetingDraft) -> str:
    """Preview of a saved meeting or appointment."""
    head = f"DRAFT INVITATION {d.ref} - NOT SENT" if d.is_meeting else f"APPOINTMENT SAVED {d.ref}"
    lines = [head, f"{d.subject} | {_span(d.start, d.end)}"]
    if d.location:
        lines.append(f"Location: {d.location}")
    for title, names in (("Required", d.required), ("Optional", d.optional), ("Rooms", d.rooms)):
        if names:
            lines.append(f"{title}: " + "; ".join(names))
    if d.text:
        lines += ["", d.text]
    return "\n".join(lines)


def free_time(result: FreeTimeResult) -> str:
    """One line per free slot, then the notes."""
    lines = [_slot(s) for s in result.slots] or ["(no common free slot in working hours)"]
    return "\n".join([*lines, *result.notes])


def _slot(s: Slot) -> str:
    return f"{fmt(s.start)}-{s.end:%H:%M}"
