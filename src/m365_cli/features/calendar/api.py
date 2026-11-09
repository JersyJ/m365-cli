from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Protocol

from m365_cli.core.paging import Page
from m365_cli.features.calendar.models import (
    Event,
    EventSummary,
    FreeTimeResult,
    MeetingDraft,
    Repeat,
    Response,
)

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

KIND = "e"
MEETING_MAIL_KIND = "m"
"""Meeting requests in the Inbox are mail items; `read()` and `respond()` accept their refs too."""


class CalendarApi(Protocol):
    """Calendar operations the CLI and MCP tools use, independent of the backend.

    Implemented by `CalendarService` (classic Outlook) and `FakeCalendarService` (fake backend).
    """

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
        ...

    def read(self, ref: str) -> Event:
        """One event; accepts e: refs and m: refs of meeting requests."""
        ...

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
        """Save a meeting (NOT sent) or, without attendees, a personal appointment."""
        ...

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

        For a meeting, send the update afterwards with `send()`.
        """
        ...

    def cleanup(self, start: datetime, end: datetime, *, apply: bool = False) -> list[EventSummary]:
        """Canceled meetings (organized by others) in the period; `apply` removes them from the calendar."""
        ...

    def send(self, ref: str) -> str:
        """Send (or re-send as update) the invitation of a meeting you organize (rate-limited, logged)."""
        ...

    def respond(self, ref: str, response: Response, *, comment: str = "") -> str:
        """Accept / tentative / decline; sends the response to the organizer (rate-limited, logged)."""
        ...

    def cancel(self, ref: str, *, comment: str = "") -> str:
        """Organizer: send a cancellation. Personal appointment: delete it."""
        ...

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
        ...


def calendar_api(ctx: AppContext) -> CalendarApi:
    """Calendar operations of the active backend."""
    if ctx.is_fake:
        from m365_cli.features.calendar.fake import FakeCalendarService

        return FakeCalendarService(ctx)
    from m365_cli.features.calendar.service import CalendarService

    return CalendarService(ctx)
