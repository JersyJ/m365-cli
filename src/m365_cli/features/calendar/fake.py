from __future__ import annotations

from datetime import datetime, timedelta

from m365_cli.backends.fake.directory import resolve_person
from m365_cli.backends.fake.state import FakeEvent, FakeState, Outgoing
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.freebusy import STATUS_NAMES, Block, BusyStatus
from m365_cli.core.paging import Page, take_page
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, truncate, with_signoff
from m365_cli.core.timeutil import now
from m365_cli.features.calendar.api import KIND, MEETING_MAIL_KIND
from m365_cli.features.calendar.models import Event, EventSummary, FreeTimeResult, MeetingDraft, Response
from m365_cli.features.calendar.rules import check_period, day_start, free_slots

_RESPONSE_STATES = {"accept": "accepted", "tentative": "tentative", "decline": "declined"}


class FakeCalendarService:
    """Calendar operations on the fake backend.

    Behaves like `CalendarService` for everything the tests and evaluations use, but never sends:
    invitations, responses and cancellations are recorded in `FakeState.outbox`.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the fake store, refs, send log and settings of `ctx`."""
        self._store = ctx.fake
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def _event(self, state: FakeState, ref: str) -> FakeEvent:
        bid = self._refs.get(ref, kinds=KIND + MEETING_MAIL_KIND)
        event_id = bid.item_id
        if bid.kind == MEETING_MAIL_KIND:
            mail = next((m for m in state.mails if m.id == bid.item_id), None)
            if mail is None or mail.event_id is None:
                raise InvalidInputError("That mail is not a meeting request.")
            event_id = mail.event_id
        for event in state.events:
            if event.id == event_id:
                return event
        raise NotFoundError("Calendar item no longer exists; list the calendar again.")

    def _summary(self, event: FakeEvent) -> EventSummary:
        return EventSummary(
            ref=self._refs.put(BackendId(KIND, event.id)),
            start=event.start,
            end=event.end,
            all_day=event.all_day,
            subject=event.subject,
            location=event.location,
            show_as=STATUS_NAMES[BusyStatus(event.show_as)],
            organizer=event.organizer,
            my_response=event.my_response if event.is_meeting else "",
            recurring=event.recurring,
            canceled=event.canceled,
            join_url=event.join_url,
        )

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
        found = sorted(
            (
                e
                for e in self._store.state.events
                if e.start < end and e.end > start and (include_canceled or not e.canceled)
            ),
            key=lambda e: e.start,
        )
        items, more = take_page(found, offset=offset, limit=limit)
        return Page(items=[self._summary(e) for e in items], offset=offset, has_more=more)

    def read(self, ref: str) -> Event:
        """One event; accepts e: refs and m: refs of meeting requests."""
        event = self._event(self._store.state, ref)
        body, truncated = truncate(clean_body(event.body), self._settings.body_max_chars)
        return Event(
            **self._summary(event).model_dump(),
            required="; ".join(event.required),
            optional="; ".join(event.optional),
            body=body,
            truncated=truncated,
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
    ) -> MeetingDraft:
        """Save a meeting (NOT sent) or, without attendees, a personal appointment."""
        check_period(start, end)
        with self._store.edit() as state:
            groups = [
                [resolve_person(state, n).name for n in group or []] for group in (required, optional, rooms)
            ]
            is_meeting = any(groups)
            text = with_signoff(body, self._settings.signoff) if is_meeting and body else body
            event = FakeEvent(
                id=state.new_id("e"),
                start=start,
                end=end,
                subject=subject,
                location=location,
                organizer=state.me.name,
                organizer_is_me=True,
                required=groups[0],
                optional=groups[1],
                rooms=groups[2],
                body=text,
                my_response="organizer" if is_meeting else "",
                is_meeting=is_meeting,
                sent=not is_meeting,
            )
            state.events.append(event)
        return MeetingDraft(
            ref=self._refs.put(BackendId(KIND, event.id)),
            subject=subject,
            start=start,
            end=end,
            location=location,
            required=groups[0],
            optional=groups[1],
            rooms=groups[2],
            text=text,
            is_meeting=is_meeting,
        )

    def send(self, ref: str) -> str:
        """Pretend to send an invitation: record it in the outbox (rate-limited and logged)."""
        self._audit.guard("cal_send")
        with self._store.edit() as state:
            event = self._event(state, ref)
            if not event.is_meeting:
                raise InvalidInputError("This is a personal appointment without attendees; nothing to send.")
            if not event.organizer_is_me:
                raise InvalidInputError("You are not the organizer of this meeting.")
            event.sent = True
            attendees = [*event.required, *event.optional, *event.rooms]
            state.outbox.append(
                Outgoing(
                    at=now(), action="cal_send", item_id=event.id, summary=event.subject, recipients=attendees
                )
            )
        summary = f"invitation '{event.subject}' to {len(attendees)} attendee(s)"
        self._audit.record("cal_send", summary)
        return f"Sent {summary}."

    def respond(self, ref: str, response: Response, *, comment: str = "") -> str:
        """Record the response in the outbox and update the event (rate-limited and logged)."""
        self._audit.guard("cal_respond")
        with self._store.edit() as state:
            event = self._event(state, ref)
            if event.organizer_is_me:
                raise InvalidInputError("You organize this meeting; use cancel instead.")
            event.my_response = _RESPONSE_STATES[response]
            event.show_as = {
                "accept": BusyStatus.BUSY,
                "tentative": BusyStatus.TENTATIVE,
                "decline": BusyStatus.FREE,
            }[response]
            state.outbox.append(
                Outgoing(
                    at=now(),
                    action="cal_respond",
                    item_id=event.id,
                    summary=f"{response}: {comment}".rstrip(": "),
                    recipients=[event.organizer],
                )
            )
        self._audit.record("cal_respond", f"{response} '{event.subject}'")
        return f"Response '{response}' sent for '{event.subject}'."

    def cancel(self, ref: str, *, comment: str = "") -> str:
        """Organizer: record a cancellation. Personal appointment: delete it."""
        with self._store.edit() as state:
            event = self._event(state, ref)
            if not event.is_meeting:
                state.events.remove(event)
                return f"Deleted appointment '{event.subject}'."
            if not event.organizer_is_me:
                raise InvalidInputError("You are not the organizer; use respond with 'decline' instead.")
            self._audit.guard("cal_cancel")
            event.canceled, event.show_as = True, BusyStatus.FREE
            state.outbox.append(
                Outgoing(
                    at=now(),
                    action="cal_cancel",
                    item_id=event.id,
                    summary=comment or event.subject,
                    recipients=event.required,
                )
            )
        self._audit.record("cal_cancel", f"'{event.subject}'")
        return f"Cancellation for '{event.subject}' sent."

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
        state = self._store.state
        first_day = day_start(day)
        blocks: list[Block] = []
        notes: list[str] = []
        for attendee in attendees:
            person = resolve_person(state, attendee)
            if person.blocks is None:
                notes.append(f"No free/busy available for {attendee}; ignored.")
                continue
            blocks += [Block(b.start, b.end, BusyStatus(b.status)) for b in person.blocks]
        if include_me:
            end = first_day + timedelta(days=days)
            blocks += [
                Block(e.start, e.end, BusyStatus(e.show_as))
                for e in state.events
                if not e.canceled and e.show_as != BusyStatus.FREE and e.start < end and e.end > first_day
            ]
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
