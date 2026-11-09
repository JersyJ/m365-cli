from __future__ import annotations

from datetime import datetime, time, timedelta
from typing import Annotated, cast

import typer

from m365_cli.core.cli_support import confirm_send, read_body, split_list
from m365_cli.core.feature import ContextProvider
from m365_cli.core.timeutil import parse_when
from m365_cli.features.calendar import render
from m365_cli.features.calendar.api import CalendarApi, calendar_api
from m365_cli.features.calendar.models import Response, ResponseChoice

Ref = Annotated[str, typer.Argument(help="Event ref e:… (or a meeting-request mail ref m:…).")]
People = Annotated[list[str] | None, typer.Option(help="Names/emails; repeat or comma-separate.")]
Yes = Annotated[bool, typer.Option("--yes", "-y", help="Do it without asking.")]


def period_end(start: datetime, end: str | None, days: int) -> datetime:
    """End of a listing period: `end` if given, else midnight `days` days after `start`."""
    return parse_when(end) if end else datetime.combine(start.date(), time()) + timedelta(days=days)


def build_cli(get_ctx: ContextProvider) -> typer.Typer:
    """`m365 calendar …` commands."""
    app = typer.Typer(help="Events, invitations, responses and free time.", no_args_is_help=True)

    def api() -> CalendarApi:
        return calendar_api(get_ctx())

    @app.command("list")
    def list_(
        start: Annotated[
            str, typer.Option("--from", help="Start: 'today', 'mon', '2026-11-10', '+2h'.")
        ] = "today",
        days: Annotated[int, typer.Option("--days", "-d", help="How many days from start.")] = 1,
        end: Annotated[str | None, typer.Option("--to", help="Explicit end (overrides --days).")] = None,
        canceled: Annotated[bool, typer.Option("--canceled", help="Include canceled events.")] = False,
        limit: Annotated[int, typer.Option("--limit", "-n")] = 50,
        offset: Annotated[int, typer.Option("--offset")] = 0,
    ) -> None:
        """List events."""
        s = parse_when(start)
        page = api().find(s, period_end(s, end, days), offset=offset, limit=limit, include_canceled=canceled)
        typer.echo(render.events(page))

    @app.command()
    def read(ref: Ref) -> None:
        """Show one event with attendees and description."""
        typer.echo(render.event(api().read(ref)))

    @app.command()
    def new(
        subject: Annotated[str, typer.Option("--subject", "-s")],
        start: Annotated[str, typer.Option("--start", help="e.g. 'tomorrow 14:00', '2026-11-12 09:30'.")],
        end: Annotated[
            str | None, typer.Option("--end", help="End time; default start + --duration.")
        ] = None,
        duration: Annotated[int, typer.Option("--duration", help="Minutes.")] = 30,
        required: People = None,
        optional: People = None,
        room: People = None,
        location: Annotated[str, typer.Option("--location", "-l")] = "",
        body: Annotated[str, typer.Option("--body", "-b", help="Description; '-' reads stdin.")] = "",
        yes: Yes = False,
        draft: Annotated[bool, typer.Option("--draft", help="Save without sending.")] = False,
    ) -> None:
        """Create a meeting (invitation preview, then confirm) or a personal appointment (no attendees)."""
        service = api()
        s = parse_when(start)
        e = parse_when(end, base=s) if end else s + timedelta(minutes=duration)
        d = service.draft(
            subject=subject,
            start=s,
            end=e,
            required=split_list(required),
            optional=split_list(optional),
            rooms=split_list(room),
            location=location,
            body=read_body(body),
        )
        if not d.is_meeting:
            typer.echo(render.draft(d))
            return
        confirm_send(
            render.draft(d),
            lambda: service.send(d.ref),
            yes=yes,
            keep_draft=draft,
            send_command=f"m365 calendar send {d.ref}",
        )

    @app.command()
    def send(ref: Ref, yes: Yes = False) -> None:
        """Send (or re-send as update) an invitation you organize."""
        if yes or typer.confirm(f"Send invitation {ref}?", default=False):
            typer.echo(api().send(ref))

    @app.command()
    def respond(
        ref: Ref,
        response: Annotated[ResponseChoice, typer.Argument(help="accept | tentative | decline")],
        comment: Annotated[str, typer.Option("--comment", "-m")] = "",
        yes: Yes = False,
    ) -> None:
        """Accept, tentatively accept or decline a meeting (sends the response)."""
        if yes or typer.confirm(f"Send '{response.value}' for {ref}?", default=False):
            typer.echo(api().respond(ref, cast(Response, response.value), comment=comment))

    @app.command()
    def cancel(
        ref: Ref, comment: Annotated[str, typer.Option("--comment", "-m")] = "", yes: Yes = False
    ) -> None:
        """Cancel a meeting you organize (notifies attendees) or delete a personal appointment."""
        if yes or typer.confirm(f"Cancel/delete {ref}?", default=False):
            typer.echo(api().cancel(ref, comment=comment))

    @app.command()
    def free(
        with_: Annotated[list[str] | None, typer.Option("--with", "-w", help="Attendees.")] = None,
        date: Annotated[str, typer.Option("--date", help="First day.")] = "today",
        days: Annotated[int, typer.Option("--days", "-d")] = 1,
        duration: Annotated[int, typer.Option("--duration", help="Minutes.")] = 30,
        limit: Annotated[int, typer.Option("--limit", "-n")] = 8,
    ) -> None:
        """Find common free slots in working hours (you + attendees)."""
        result = api().find_free_time(
            split_list(with_), day=parse_when(date), days=days, duration_minutes=duration, limit=limit
        )
        typer.echo(render.free_time(result))

    return app
