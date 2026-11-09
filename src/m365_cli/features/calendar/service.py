from __future__ import annotations

import contextlib
import re
from datetime import datetime, timedelta

import pywintypes

from m365_cli.backends.outlook import Com, Outlook, iterate
from m365_cli.backends.outlook.constants import (
    OL_DISCARD,
    OlItemType,
    OlMeetingResponse,
    OlMeetingStatus,
    OlRecipientType,
    OlResponseStatus,
)
from m365_cli.backends.outlook.freebusy import own_calendar_blocks, recipient_blocks
from m365_cli.backends.outlook.items import get
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.freebusy import STATUS_NAMES, Block, BusyStatus
from m365_cli.core.paging import Page, take_page
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, truncate, with_signoff
from m365_cli.core.timeutil import from_com, now, to_com
from m365_cli.features.calendar.api import KIND, MEETING_MAIL_KIND
from m365_cli.features.calendar.models import Event, EventSummary, FreeTimeResult, MeetingDraft, Response
from m365_cli.features.calendar.rules import check_period, day_start, free_slots

_OCCURRENCE_SEP = "|"
_JOIN_URL = re.compile(r"https://teams\.microsoft\.com/l/meetup-join/[^\s<>\"']+")
_RESPONSE_NAMES = {
    OlResponseStatus.ORGANIZED: "organizer",
    OlResponseStatus.TENTATIVE: "tentative",
    OlResponseStatus.ACCEPTED: "accepted",
    OlResponseStatus.DECLINED: "declined",
    OlResponseStatus.NOT_RESPONDED: "not responded",
}
_RESPONSES = {
    "accept": OlMeetingResponse.ACCEPTED,
    "tentative": OlMeetingResponse.TENTATIVE,
    "decline": OlMeetingResponse.DECLINED,
}
_FREEBUSY_MINUTES = 15
_NO_UI = True


class CalendarService:
    """Calendar operations on classic Outlook.

    List/read events, draft & send invitations, respond, cancel, and find common free time. Like
    mail, anything that notifies other people is a separate, explicit step.

    Occurrences of a recurring meeting share the series EntryID, so their ref also encodes the start;
    moved occurrences (Outlook "exceptions") are found by their actual start.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook session, refs, send log and settings of `ctx`."""
        self._session = ctx.outlook
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def _ref_for(self, ol: Outlook, item: Com) -> str:
        item_id = item.EntryID
        if get(item, "IsRecurring", False):
            item_id += _OCCURRENCE_SEP + _start(item).isoformat()
        return self._refs.put(BackendId(KIND, item_id, ol.default_store_id))

    @staticmethod
    def _appointment(ol: Outlook, bid: BackendId) -> Com:
        if bid.kind == MEETING_MAIL_KIND:
            appointment = ol.item(bid).GetAssociatedAppointment(True)
            if appointment is None:
                raise InvalidInputError("That mail is not a meeting request.")
            return appointment
        entry_id, _, occurrence = bid.item_id.partition(_OCCURRENCE_SEP)
        item = ol.item(BackendId(bid.kind, entry_id, bid.container_id))
        return _occurrence(ol, item, datetime.fromisoformat(occurrence)) if occurrence else item

    def find(
        self,
        start: datetime,
        end: datetime,
        *,
        offset: int = 0,
        limit: int = 50,
        include_canceled: bool = False,
    ) -> Page[EventSummary]:
        """One page of events overlapping the period, earliest first."""
        check_period(start, end)

        def run(ol: Outlook) -> Page[EventSummary]:
            summaries = (self._summary(ol, item) for item in iterate(ol.calendar_between(start, end)))
            visible = (s for s in summaries if include_canceled or not s.canceled)
            items, more = take_page(visible, offset=offset, limit=limit)
            return Page(items=items, offset=offset, has_more=more)

        return self._session.run(run)

    def _summary(self, ol: Outlook, item: Com) -> EventSummary:
        body = clean_body(str(get(item, "Body", "") or ""))
        join = _JOIN_URL.search(body)
        status = get(item, "MeetingStatus", 0)
        is_meeting = status != OlMeetingStatus.NON_MEETING
        return EventSummary(
            ref=self._ref_for(ol, item),
            start=_start(item),
            end=from_com(item.End) or _start(item),
            all_day=bool(get(item, "AllDayEvent", False)),
            subject=str(get(item, "Subject", "") or "(no subject)"),
            location=str(get(item, "Location", "") or ""),
            show_as=STATUS_NAMES.get(BusyStatus(get(item, "BusyStatus", BusyStatus.BUSY)), "busy"),
            organizer=str(get(item, "Organizer", "") or ""),
            my_response=_RESPONSE_NAMES.get(get(item, "ResponseStatus", 0), "") if is_meeting else "",
            recurring=bool(get(item, "IsRecurring", False)),
            canceled=status in (OlMeetingStatus.CANCELED, OlMeetingStatus.RECEIVED_AND_CANCELED),
            join_url=join.group(0) if join else "",
        )

    def read(self, ref: str) -> Event:
        """One event; accepts e: refs and m: refs of meeting requests."""
        bid = self._refs.get(ref, kinds=KIND + MEETING_MAIL_KIND)

        def run(ol: Outlook) -> Event:
            item = self._appointment(ol, bid)
            body, truncated = truncate(
                clean_body(str(get(item, "Body", "") or "")), self._settings.body_max_chars
            )
            return Event(
                **self._summary(ol, item).model_dump(),
                required=str(get(item, "RequiredAttendees", "") or ""),
                optional=str(get(item, "OptionalAttendees", "") or ""),
                categories=str(get(item, "Categories", "") or ""),
                body=body,
                truncated=truncated,
            )

        return self._session.run(run)

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
    ) -> MeetingDraft:
        """Save a meeting (NOT sent) or, without attendees, a personal appointment.

        Raises:
            InvalidInputError: End is not after start.
            UnresolvedRecipientsError: An attendee is unknown; nothing is saved.
        """
        check_period(start, end)
        groups = (
            (required, OlRecipientType.TO),
            (optional, OlRecipientType.CC),
            (rooms, OlRecipientType.BCC),
        )
        attendees = [(address, rtype) for group, rtype in groups for address in group or []]
        text = with_signoff(body, self._settings.signoff) if attendees and body else body

        def run(ol: Outlook) -> MeetingDraft:
            item = ol.app.CreateItem(OlItemType.APPOINTMENT)
            item.Subject, item.Start, item.End = subject, to_com(start), to_com(end)
            item.Location, item.Body = location, text
            if attendees:
                item.MeetingStatus = OlMeetingStatus.MEETING
                for address, rtype in attendees:
                    item.Recipients.Add(address).Type = rtype
                try:
                    ol.resolve_all(item)
                except Exception:
                    item.Close(OL_DISCARD)
                    raise
            item.Save()
            names = {t: [r.Name for r in item.Recipients if r.Type == t] for t in OlRecipientType}
            return MeetingDraft(
                ref=self._ref_for(ol, item),
                subject=subject,
                start=start,
                end=end,
                location=location,
                required=names[OlRecipientType.TO],
                optional=names[OlRecipientType.CC],
                rooms=names[OlRecipientType.BCC],
                text=text,
                is_meeting=bool(attendees),
            )

        return self._session.run(run)

    def send(self, ref: str) -> str:
        """Send (or re-send as update) the invitation of a meeting you organize.

        Raises:
            InvalidInputError: No attendees, or you are not the organizer.
            RateLimitError: Too many sends in the last minute.
        """
        bid = self._refs.get(ref, kinds=KIND)
        self._audit.guard("cal_send")

        def run(ol: Outlook) -> str:
            item = self._appointment(ol, bid)
            if item.MeetingStatus == OlMeetingStatus.NON_MEETING:
                raise InvalidInputError("This is a personal appointment without attendees; nothing to send.")
            if get(item, "ResponseStatus", 0) not in (OlResponseStatus.ORGANIZED, OlResponseStatus.NONE):
                raise InvalidInputError("You are not the organizer of this meeting.")
            ol.resolve_all(item)
            item.Send()
            return f"invitation '{item.Subject}' to {item.Recipients.Count} attendee(s)"

        summary = self._session.run(run)
        self._audit.record("cal_send", summary)
        return f"Sent {summary}."

    def respond(self, ref: str, response: Response, *, comment: str = "") -> str:
        """Accept / tentative / decline; sends the response to the organizer.

        Raises:
            InvalidInputError: You organize this meeting.
            RateLimitError: Too many sends in the last minute.
        """
        bid = self._refs.get(ref, kinds=KIND + MEETING_MAIL_KIND)
        self._audit.guard("cal_respond")

        def run(ol: Outlook) -> tuple[str, bool]:
            appointment = self._appointment(ol, bid)
            if get(appointment, "ResponseStatus", 0) == OlResponseStatus.ORGANIZED:
                raise InvalidInputError("You organize this meeting; use cancel instead.")
            reply = appointment.Respond(_RESPONSES[response], _NO_UI)
            if reply is None:
                appointment.Save()
                return str(appointment.Subject), False
            if comment:
                reply.Body = comment
            reply.Send()
            return str(appointment.Subject), True

        subject, sent = self._session.run(run)
        if not sent:
            return f"Marked '{subject}' as {response} (no response requested)."
        self._audit.record("cal_respond", f"{response} '{subject}'")
        return f"Response '{response}' sent for '{subject}'."

    def cancel(self, ref: str, *, comment: str = "") -> str:
        """Organizer: send a cancellation. Personal appointment: delete it.

        Raises:
            InvalidInputError: You are an attendee, not the organizer.
            RateLimitError: Too many sends in the last minute.
        """
        bid = self._refs.get(ref, kinds=KIND)

        def run(ol: Outlook) -> tuple[str, bool]:
            item = self._appointment(ol, bid)
            subject = str(item.Subject)
            if item.MeetingStatus == OlMeetingStatus.NON_MEETING:
                item.Delete()
                return subject, False
            if get(item, "ResponseStatus", 0) != OlResponseStatus.ORGANIZED:
                raise InvalidInputError("You are not the organizer; use respond with 'decline' instead.")
            self._audit.guard("cal_cancel")
            item.MeetingStatus = OlMeetingStatus.CANCELED
            if comment:
                item.Body = f"{comment}\n\n{item.Body}"
            item.Send()
            return subject, True

        subject, sent = self._session.run(run)
        if not sent:
            return f"Deleted appointment '{subject}'."
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
        """Common free slots in working hours for you and the attendees."""
        first_day = day_start(day)

        def run(ol: Outlook) -> FreeTimeResult:
            blocks: list[Block] = []
            notes: list[str] = []
            for attendee in attendees:
                found = recipient_blocks(ol.recipient(attendee), first_day, _FREEBUSY_MINUTES, days)
                if found is None:
                    notes.append(f"No free/busy available for {attendee}; ignored.")
                else:
                    blocks += found
            if include_me:
                blocks += own_calendar_blocks(ol, first_day, first_day + timedelta(days=days))
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

        return self._session.run(run)


def _start(item: Com) -> datetime:
    start = from_com(item.Start)
    if start is None:
        raise InvalidInputError("Calendar item has no start time.")
    return start


def _occurrence(ol: Outlook, series: Com, start: datetime) -> Com:
    """One occurrence of a recurring series, including occurrences that were moved (exceptions)."""
    with contextlib.suppress(pywintypes.com_error):
        return series.GetRecurrencePattern().GetOccurrence(to_com(start))
    for item in iterate(ol.calendar_between(start, start + timedelta(minutes=1))):
        if item.EntryID == series.EntryID and from_com(item.Start) == start:
            return item
    raise NotFoundError("That occurrence no longer exists; list the calendar again.")
