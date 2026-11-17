from __future__ import annotations

import re
import urllib.parse
from datetime import datetime, timedelta
from typing import Any

from m365_cli.backends.outlook_web.session import ews_id, from_wall, rest_id, wall
from m365_cli.core.context import AppContext
from m365_cli.core.errors import (
    InvalidInputError,
    M365Error,
    NotFoundError,
    UnresolvedRecipientsError,
)
from m365_cli.core.freebusy import Block, parse_freebusy
from m365_cli.core.paging import Page, take_page
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, truncate, with_signoff
from m365_cli.core.timeutil import now
from m365_cli.features.calendar.api import KIND, MEETING_MAIL_KIND
from m365_cli.features.calendar.models import (
    Event,
    EventSummary,
    FreeTimeResult,
    MeetingDraft,
    Repeat,
    Response,
)
from m365_cli.features.calendar.rules import check_period, check_repeat, day_start, free_slots

_FIELDS = (
    "Id,Subject,Start,End,IsAllDay,Location,ShowAs,Organizer,ResponseStatus,Type,IsCancelled,"
    "IsOrganizer,Attendees,OnlineMeetingUrl,OnlineMeeting,ResponseRequested,SeriesMasterId"
)
_SHOW_AS = {
    "free": "free",
    "tentative": "tentative",
    "busy": "busy",
    "oof": "out of office",
    "workingelsewhere": "working elsewhere",
}
_RESPONSE_NAMES = {
    "organizer": "organizer",
    "accepted": "accepted",
    "tentativelyaccepted": "tentative",
    "declined": "declined",
    "notresponded": "not responded",
}
_RESPONSES = {"accept": "accept", "tentative": "tentativelyaccept", "decline": "decline"}
_ATTENDEE_FIELDS = {"required": "RequiredAttendees", "optional": "OptionalAttendees", "resource": "Resources"}
_WEEKDAYS = {
    "mon": "Monday",
    "tue": "Tuesday",
    "wed": "Wednesday",
    "thu": "Thursday",
    "fri": "Friday",
    "sat": "Saturday",
    "sun": "Sunday",
}
_JOIN_URL = re.compile(r"https://teams\.microsoft\.com/l/meetup-join/[^\s<>\"']+")
_FREEBUSY_MINUTES = 15
_MAX_EVENTS = 2000


class WebCalendarService:
    """Calendar operations on Outlook on the web.

    Times are wall-clock times in the mailbox time zone. Anything that notifies other people is a
    separate, explicit step: meetings are saved and changed through service.svc without sending,
    and only `send`, `respond` and `cancel` notify.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook web session, refs, send log and settings of `ctx`."""
        self._web = ctx.outlook_web
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def _rest(self, path: str, **kwargs: Any) -> Any:
        return self._web.rest(path, timezone=self._web.timezone(), **kwargs)

    def _event_id(self, ref: str) -> tuple[str, str | None]:
        """(event id, owner mailbox) of an e: ref, or of the event of an m: meeting-request ref."""
        bid = self._refs.get(ref, kinds=KIND + MEETING_MAIL_KIND)
        owner = bid.container_id or None
        if bid.kind != MEETING_MAIL_KIND:
            return bid.item_id, owner
        answer = self._web.svc(
            "GetItem",
            {
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "AdditionalProperties": [
                        {"__type": "PropertyUri:#Exchange", "FieldURI": "AssociatedCalendarItemId"}
                    ],
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": ews_id(bid.item_id)}],
            },
            mailbox=owner,
        )
        found = ((answer.get("Items") or [{}])[0].get("AssociatedCalendarItemId") or {}).get("Id")
        if not found:
            raise InvalidInputError("That mail is not a meeting request.")
        return rest_id(str(found)), owner

    def _get(self, event_id: str, owner: str | None, *, body: bool = False) -> dict[str, Any]:
        select = _FIELDS + (",Body,Categories" if body else "")
        return self._rest(f"/events/{event_id}?$select={select}", mailbox=owner, text=True) or {}

    def _summary(self, ev: dict[str, Any], owner: str | None) -> EventSummary:
        start = from_wall(ev.get("Start")) or now()
        is_meeting = bool(ev.get("Attendees")) or not ev.get("IsOrganizer", True)
        response = str((ev.get("ResponseStatus") or {}).get("Response") or "").lower()
        online = ev.get("OnlineMeeting") or {}
        body = str((ev.get("Body") or {}).get("Content") or "")
        join = online.get("JoinUrl") or ev.get("OnlineMeetingUrl") or ""
        if not join and (m := _JOIN_URL.search(body)):
            join = m.group(0)
        return EventSummary(
            ref=self._refs.put(BackendId(KIND, str(ev["Id"]), owner or "")),
            start=start,
            end=from_wall(ev.get("End")) or start,
            all_day=bool(ev.get("IsAllDay")),
            subject=str(ev.get("Subject") or "(no subject)"),
            location=str((ev.get("Location") or {}).get("DisplayName") or ""),
            show_as=_SHOW_AS.get(str(ev.get("ShowAs") or "busy").lower(), "busy"),
            organizer=str(((ev.get("Organizer") or {}).get("EmailAddress") or {}).get("Name") or ""),
            my_response=_RESPONSE_NAMES.get(response, "") if is_meeting else "",
            recurring=str(ev.get("Type") or "") in ("Occurrence", "Exception", "SeriesMaster"),
            canceled=bool(ev.get("IsCancelled")),
            join_url=str(join),
        )

    def _view(self, start: datetime, end: datetime, owner: str | None = None) -> list[dict[str, Any]]:
        """Events (occurrences expanded) overlapping the period, earliest first."""
        params = urllib.parse.urlencode(
            {
                "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%S"),
                "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%S"),
                "$select": _FIELDS,
                "$orderby": "Start/DateTime",
                "$top": "200",
            },
            safe="/$,",
        )
        try:
            return self._web.rest_pages(
                f"/calendarview?{params}", max_items=_MAX_EVENTS, mailbox=owner, timezone=self._web.timezone()
            )
        except M365Error as exc:
            if owner and not isinstance(exc, InvalidInputError):
                raise InvalidInputError(
                    f"{owner}'s calendar is not shared with you; use cal_find_free_time or people_lookup for "
                    "their free/busy instead."
                ) from exc
            raise

    def find(
        self,
        start: datetime,
        end: datetime,
        *,
        offset: int = 0,
        limit: int = 50,
        include_canceled: bool = False,
        owner: str | None = None,
    ) -> Page[EventSummary]:
        """One page of events overlapping the period, earliest first; `owner` = a shared calendar."""
        check_period(start, end)
        box = self._web.mailbox(owner) if owner else None
        summaries = (self._summary(ev, box) for ev in self._view(start, end, box))
        visible = (s for s in summaries if include_canceled or not s.canceled)
        items, more = take_page(visible, offset=offset, limit=limit)
        return Page(items=items, offset=offset, has_more=more)

    def read(self, ref: str) -> Event:
        """One event; accepts e: refs and m: refs of meeting requests."""
        event_id, owner = self._event_id(ref)
        ev = self._get(event_id, owner, body=True)
        body, truncated = truncate(
            clean_body(str((ev.get("Body") or {}).get("Content") or "")), self._settings.body_max_chars
        )
        return Event(
            **self._summary(ev, owner).model_dump(),
            required=_names(ev, "required"),
            optional=_names(ev, "optional"),
            categories=", ".join(ev.get("Categories") or []),
            body=body,
            truncated=truncated,
        )

    def _attendees(self, values: list[str] | None) -> list[tuple[str, str]]:
        resolved, missing = [], []
        for value in values or []:
            found = self._web.resolve(value)
            if found is None:
                missing.append(value)
            else:
                resolved.append(found)
        if missing:
            raise UnresolvedRecipientsError(missing)
        return resolved

    def _set_attendees(self, event_id: str, groups: dict[str, list[tuple[str, str]]], **fields: Any) -> None:
        """Save attendees (and other fields) of an event without notifying anyone (service.svc)."""
        updates = [
            _set_field(_ATTENDEE_FIELDS[kind], _attendee_list(people)) for kind, people in groups.items()
        ]
        updates += [_set_field(name, value) for name, value in fields.items()]
        if not updates:
            return
        self._web.svc(
            "UpdateItem",
            {
                "ConflictResolution": "AlwaysOverwrite",
                "SendCalendarInvitationsOrCancellations": "SendToNone",
                "ItemChanges": [
                    {
                        "__type": "ItemChange:#Exchange",
                        "ItemId": {"__type": "ItemId:#Exchange", "Id": ews_id(event_id)},
                        "Updates": updates,
                    }
                ],
            },
        )

    def draft(
        self,
        *,
        subject: str,
        start: datetime,
        end: datetime,
        required: list[str] | None = None,
        optional: list[str] | None = None,
        rooms: list[str] | None = None,
        location: str = "",
        body: str = "",
        repeat: Repeat | None = None,
    ) -> MeetingDraft:
        """Save a meeting (NOT sent) or, without attendees, a personal appointment.

        The event is created without attendees, then the attendees are added without notifying anyone.

        Raises:
            InvalidInputError: End is not after start.
            UnresolvedRecipientsError: An attendee is unknown; nothing is saved.
        """
        check_period(start, end)
        weekdays = check_repeat(repeat, start) if repeat else []
        groups = {
            "required": self._attendees(required),
            "optional": self._attendees(optional),
            "resource": self._attendees(rooms),
        }
        groups = {k: v for k, v in groups.items() if v}
        text = with_signoff(body, self._settings.signoff) if groups and body else body
        zone = self._web.timezone()
        event: dict[str, Any] = {
            "Subject": subject,
            "Start": wall(start, zone),
            "End": wall(end, zone),
            "Location": {"DisplayName": location},
            "Body": {"ContentType": "Text", "Content": text},
        }
        if repeat:
            event["Recurrence"] = _recurrence(repeat, weekdays, start, zone)
        created = self._rest("/events", method="POST", body=event) or {}
        event_id = str(created["Id"])
        try:
            self._set_attendees(event_id, groups)
        except Exception:
            self._rest(f"/events/{event_id}", method="DELETE")
            raise
        return MeetingDraft(
            ref=self._refs.put(BackendId(KIND, event_id, "")),
            subject=subject,
            start=start,
            end=end,
            location=location,
            required=[n for n, _ in groups.get("required", [])],
            optional=[n for n, _ in groups.get("optional", [])],
            rooms=[n for n, _ in groups.get("resource", [])],
            text=text,
            is_meeting=bool(groups),
            repeat=repeat,
        )

    def update(
        self,
        ref: str,
        *,
        start: datetime | None = None,
        end: datetime | None = None,
        subject: str | None = None,
        location: str | None = None,
        add_required: list[str] | None = None,
        add_optional: list[str] | None = None,
        remove: list[str] | None = None,
    ) -> MeetingDraft:
        """Change a meeting you organize (or a personal appointment) and save it; attendees are NOT notified.

        A new start without an end keeps the duration. For a meeting, send the update with `send()`.

        Raises:
            InvalidInputError: Nothing to change, end before start, or you are not the organizer.
            NotFoundError: An attendee to remove is not on the meeting.
            UnresolvedRecipientsError: A new attendee is unknown; nothing is saved.
        """
        changes = (start, end, subject, location, add_required, add_optional, remove)
        if all(c is None for c in changes):
            raise InvalidInputError("Nothing to update.")
        event_id, owner = self._event_id(ref)
        ev = self._get(event_id, owner)
        if not ev.get("IsOrganizer", True):
            raise InvalidInputError("You are not the organizer; ask the organizer or use respond.")
        old_start = from_wall(ev.get("Start")) or now()
        old_end = from_wall(ev.get("End")) or old_start
        new_start = start or old_start
        new_end = end or (new_start + (old_end - old_start) if start else old_end)
        check_period(new_start, new_end)
        groups = self._changed_attendees(ev, add_required, add_optional, remove)
        zone = self._web.timezone()
        if groups or ev.get("Attendees"):
            fields: dict[str, Any] = {}
            if start or end:
                fields |= {
                    "Start": new_start.strftime("%Y-%m-%dT%H:%M:%S"),
                    "End": new_end.strftime("%Y-%m-%dT%H:%M:%S"),
                }
            if subject is not None:
                fields["Subject"] = subject
            if location is not None:
                fields["Location"] = location
            self._set_attendees(event_id, groups, **fields)
        else:
            patch: dict[str, Any] = {}
            if start or end:
                patch |= {"Start": wall(new_start, zone), "End": wall(new_end, zone)}
            if subject is not None:
                patch["Subject"] = subject
            if location is not None:
                patch["Location"] = {"DisplayName": location}
            self._rest(f"/events/{event_id}", method="PATCH", mailbox=owner, body=patch)
        saved = self._get(event_id, owner)
        return MeetingDraft(
            ref=self._refs.put(BackendId(KIND, event_id, owner or "")),
            subject=str(saved.get("Subject") or ""),
            start=from_wall(saved.get("Start")) or new_start,
            end=from_wall(saved.get("End")) or new_end,
            location=str((saved.get("Location") or {}).get("DisplayName") or ""),
            required=[_name(a) for a in saved.get("Attendees") or [] if _kind(a) == "required"],
            optional=[_name(a) for a in saved.get("Attendees") or [] if _kind(a) == "optional"],
            rooms=[_name(a) for a in saved.get("Attendees") or [] if _kind(a) == "resource"],
            is_meeting=bool(saved.get("Attendees")),
            is_update=True,
        )

    def _changed_attendees(
        self,
        ev: dict[str, Any],
        add_required: list[str] | None,
        add_optional: list[str] | None,
        remove: list[str] | None,
    ) -> dict[str, list[tuple[str, str]]]:
        """New full attendee lists of the kinds that change (required, optional, resource).

        Raises:
            NotFoundError: An attendee to remove is not on the meeting.
            UnresolvedRecipientsError: A new attendee is unknown.
        """
        current = {
            kind: [(_name(a), _address(a)) for a in ev.get("Attendees") or [] if _kind(a) == kind]
            for kind in _ATTENDEE_FIELDS
        }
        changed: set[str] = set()
        for query in remove or []:
            key = query.strip().lower()
            gone = {
                k: [p for p in v if key in p[0].lower() or key in p[1].lower()] for k, v in current.items()
            }
            if not any(gone.values()):
                raise NotFoundError(f"{query} is not an attendee of this meeting.")
            for kind, people in gone.items():
                if people:
                    current[kind] = [p for p in current[kind] if p not in people]
                    changed.add(kind)
        for kind, added in (("required", add_required), ("optional", add_optional)):
            if added:
                current[kind] = [*current[kind], *self._attendees(added)]
                changed.add(kind)
        return {kind: current[kind] for kind in changed}

    def _was_sent(self, event_id: str) -> bool:
        """True when the invitation of a meeting you organize has been sent at least once."""
        answer = self._web.svc(
            "GetItem",
            {
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "AdditionalProperties": [
                        {"__type": "PropertyUri:#Exchange", "FieldURI": "MeetingRequestWasSent"}
                    ],
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": ews_id(event_id)}],
            },
        )
        return bool((answer.get("Items") or [{}])[0].get("MeetingRequestWasSent", True))

    def _delete_silently(self, *event_ids: str) -> None:
        """Move events to Deleted Items without sending cancellations."""
        self._web.svc(
            "DeleteItem",
            {
                "DeleteType": "MoveToDeletedItems",
                "SendMeetingCancellations": "SendToNone",
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": ews_id(i)} for i in event_ids],
            },
        )

    def cleanup(self, start: datetime, end: datetime, *, apply: bool = False) -> list[EventSummary]:
        """Canceled meetings (organized by others) in the period; with `apply`, remove them from the calendar.

        Removing a canceled meeting notifies nobody.
        """
        check_period(start, end)
        targets = [ev for ev in self._view(start, end) if ev.get("IsCancelled") and not ev.get("IsOrganizer")]
        summaries = [self._summary(ev, None) for ev in targets]
        if apply and targets:
            self._delete_silently(*(str(ev["Id"]) for ev in targets))
        return summaries

    def send(self, ref: str) -> str:
        """Send (or re-send as update) the invitation of a meeting you organize.

        Raises:
            InvalidInputError: No attendees, or you are not the organizer.
            RateLimitError: Too many sends in the last minute.
        """
        event_id, owner = self._event_id(ref)
        ev = self._get(event_id, owner)
        if not ev.get("Attendees"):
            raise InvalidInputError("This is a personal appointment without attendees; nothing to send.")
        if not ev.get("IsOrganizer", True):
            raise InvalidInputError("You are not the organizer of this meeting.")
        self._audit.guard("cal_send")
        subject = str(ev.get("Subject") or "")
        self._web.svc(
            "UpdateItem",
            {
                "ConflictResolution": "AlwaysOverwrite",
                "SendCalendarInvitationsOrCancellations": "SendToAllAndSaveCopy",
                "ItemChanges": [
                    {
                        "__type": "ItemChange:#Exchange",
                        "ItemId": {"__type": "ItemId:#Exchange", "Id": ews_id(event_id)},
                        "Updates": [_set_field("Subject", subject)],
                    }
                ],
            },
        )
        summary = f"invitation '{subject}' to {len(ev['Attendees'])} attendee(s)"
        self._audit.record("cal_send", summary)
        return f"Sent {summary}."

    def respond(self, ref: str, response: Response, *, comment: str = "", series: bool = False) -> str:
        """Accept / tentative / decline; sends the response to the organizer when one is requested.

        `series` answers the whole recurring series (its master) instead of one occurrence.

        Raises:
            InvalidInputError: You organize this meeting, or `series` for a meeting that does not repeat.
            RateLimitError: Too many sends in the last minute.
        """
        event_id, owner = self._event_id(ref)
        ev = self._get(event_id, owner)
        if series and str(ev.get("Type") or "") != "SeriesMaster":
            master = ev.get("SeriesMasterId")
            if not master:
                raise InvalidInputError("This meeting is not part of a series.")
            event_id = str(master)
            ev = self._get(event_id, owner)
        if ev.get("IsOrganizer"):
            raise InvalidInputError("You organize this meeting; use cancel instead.")
        requested = bool(ev.get("ResponseRequested", True))
        if requested:
            self._audit.guard("cal_respond")
        self._rest(
            f"/events/{event_id}/{_RESPONSES[response]}",
            method="POST",
            mailbox=owner,
            body={"Comment": comment, "SendResponse": True} if requested else {"SendResponse": False},
        )
        subject = str(ev.get("Subject") or "") + (" (whole series)" if series else "")
        if not requested:
            return f"Marked '{subject}' as {response} (no response requested)."
        self._audit.record("cal_respond", f"{response} '{subject}'")
        return f"Response '{response}' sent for '{subject}'."

    def cancel(self, ref: str, *, comment: str = "") -> str:
        """Organizer: send a cancellation. Personal appointment: delete it.

        Raises:
            InvalidInputError: You are an attendee, not the organizer.
            RateLimitError: Too many sends in the last minute.
        """
        event_id, owner = self._event_id(ref)
        ev = self._get(event_id, owner)
        subject = str(ev.get("Subject") or "")
        if not ev.get("Attendees") and ev.get("IsOrganizer", True):
            self._rest(f"/events/{event_id}", method="DELETE", mailbox=owner)
            return f"Deleted appointment '{subject}'."
        if not ev.get("IsOrganizer"):
            raise InvalidInputError("You are not the organizer; use respond with 'decline' instead.")
        if not self._was_sent(event_id):
            self._delete_silently(event_id)
            return f"Deleted '{subject}'; its invitation was never sent, so nobody is notified."
        self._audit.guard("cal_cancel")
        self._rest(f"/events/{event_id}/cancel", method="POST", mailbox=owner, body={"Comment": comment})
        self._audit.record("cal_cancel", f"'{subject}'")
        return f"Cancellation for '{subject}' sent."

    def find_free_time(
        self,
        attendees: list[str],
        *,
        day: datetime,
        days: int = 1,
        duration_minutes: int = 30,
        limit: int = 8,
        include_me: bool = True,
    ) -> FreeTimeResult:
        """Common free slots in working hours for you and the attendees (Exchange free/busy)."""
        first_day = day_start(day)
        people = [address for _, address in self._attendees(attendees)]
        if include_me:
            people.append(self._web.me()[1])
        blocks, notes = self._free_busy(people, first_day, days, _FREEBUSY_MINUTES)
        slots = free_slots(
            blocks,
            first_day=first_day,
            days=days,
            duration_minutes=duration_minutes,
            work_hours=self._settings.work_hours,
            earliest=now(),
            limit=limit,
        )
        return FreeTimeResult(slots=slots, notes=notes)

    def _free_busy(
        self, addresses: list[str], first_day: datetime, days: int, minutes: int
    ) -> tuple[list[Block], list[str]]:
        """Free/busy blocks of everyone (from `first_day` midnight, for `days` days), plus notes."""
        blocks: list[Block] = []
        notes: list[str] = []
        for address, view in self._web.schedules(
            addresses, first_day, first_day + timedelta(days=days), minutes
        ):
            if view is None:
                notes.append(f"No free/busy available for {address}; ignored.")
            else:
                blocks += parse_freebusy(view, first_day, minutes)
        return blocks, notes


def _kind(attendee: dict[str, Any]) -> str:
    return str(attendee.get("Type") or "Required").lower()


def _name(attendee: dict[str, Any]) -> str:
    address = attendee.get("EmailAddress") or {}
    return str(address.get("Name") or address.get("Address") or "")


def _address(attendee: dict[str, Any]) -> str:
    return str((attendee.get("EmailAddress") or {}).get("Address") or "")


def _names(ev: dict[str, Any], kind: str) -> str:
    return "; ".join(_name(a) for a in ev.get("Attendees") or [] if _kind(a) == kind)


def _attendee_list(people: list[tuple[str, str]]) -> list[dict[str, Any]]:
    return [
        {
            "__type": "AttendeeType:#Exchange",
            "Mailbox": {"__type": "EmailAddress:#Exchange", "Name": name, "EmailAddress": address},
        }
        for name, address in people
    ]


def _set_field(name: str, value: Any) -> dict[str, Any]:
    """One service.svc field change; the location is set through its 'EnhancedLocation' field."""
    if name == "Location":
        return {
            "__type": "SetItemField:#Exchange",
            "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": "EnhancedLocation"},
            "Item": {
                "__type": "CalendarItem:#Exchange",
                "Location": {"__type": "EnhancedLocation:#Exchange", "DisplayName": value},
            },
        }
    return {
        "__type": "SetItemField:#Exchange",
        "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": name},
        "Item": {"__type": "CalendarItem:#Exchange", name: value},
    }


def _recurrence(repeat: Repeat, weekdays: list[str], start: datetime, zone: str) -> dict[str, Any]:
    """REST recurrence of a repeat rule (the pattern keeps the start's time of day, in the mailbox zone)."""
    pattern: dict[str, Any] = {
        "Type": {"day": "Daily", "week": "Weekly", "month": "AbsoluteMonthly"}[repeat.every],
        "Interval": repeat.interval,
    }
    if weekdays:
        pattern |= {"DaysOfWeek": [_WEEKDAYS[d] for d in weekdays], "FirstDayOfWeek": "Monday"}
    if repeat.every == "month":
        pattern["DayOfMonth"] = start.day
    rng: dict[str, Any] = {"StartDate": start.date().isoformat(), "RecurrenceTimeZone": zone}
    if repeat.until is not None:
        rng |= {"Type": "EndDate", "EndDate": repeat.until.isoformat()}
    elif repeat.count is not None:
        rng |= {"Type": "Numbered", "NumberOfOccurrences": repeat.count}
    else:
        rng["Type"] = "NoEnd"
    return {"Pattern": pattern, "Range": rng}
