from __future__ import annotations

import json
from collections.abc import Callable
from typing import Annotated

import typer

from m365_cli.core.cli_support import confirm_send, read_body
from m365_cli.core.feature import ContextProvider
from m365_cli.core.markup import BodyFormat
from m365_cli.core.timeutil import parse_since, parse_when
from m365_cli.features.teams import render
from m365_cli.features.teams.api import TeamsApi, teams_api
from m365_cli.features.teams.models import ChatKind

ChatArg = Annotated[
    str, typer.Argument(help="Chat ref (c:…) or name (chat title, or the person of a 1:1 chat).")
]
Yes = Annotated[bool, typer.Option("--yes", "-y", help="Send without asking.")]
KeepDraft = Annotated[bool, typer.Option("--draft", help="Only save the draft, do not send.")]
At = Annotated[str | None, typer.Option("--at", help="Schedule, e.g. 'tomorrow 9:00'.")]


def build_cli(get_ctx: ContextProvider) -> typer.Typer:
    """`m365 teams …` commands."""
    app = typer.Typer(help="Teams chats and channels (in the dedicated Teams browser).", no_args_is_help=True)

    def api() -> TeamsApi:
        return teams_api(get_ctx())

    _read_commands(app, api, get_ctx)
    _write_commands(app, api)
    return app


def _read_commands(app: typer.Typer, api: Callable[[], TeamsApi], get_ctx: ContextProvider) -> None:
    """Commands that only read: list, digest, read, search, files, status, browser."""

    @app.command("list")
    def list_(
        query: Annotated[str | None, typer.Argument(help="Words in the title or person's name.")] = None,
        kind: Annotated[ChatKind | None, typer.Option("--kind", "-k")] = None,
        unread: Annotated[bool, typer.Option("--unread", "-u")] = False,
        limit: Annotated[int, typer.Option("--limit", "-n")] = 15,
        offset: Annotated[int, typer.Option("--offset")] = 0,
    ) -> None:
        """List chats, newest activity first."""
        page = api().chats(query=query, kind=kind, unread_only=unread, offset=offset, limit=limit)
        typer.echo(render.chats(page))

    @app.command()
    def digest(
        all_: Annotated[bool, typer.Option("--all", help="Include read mentions too.")] = False,
        limit: Annotated[int, typer.Option("--limit", "-n")] = 15,
    ) -> None:
        """What needs you on Teams: mentions/replies and unread conversations."""
        typer.echo(render.digest(api().digest(unread_only=not all_, limit=limit)))

    @app.command()
    def read(
        chat: ChatArg,
        limit: Annotated[int, typer.Option("--limit", "-n")] = 20,
        since: Annotated[str | None, typer.Option("--since", help="e.g. '2d', 'today'.")] = None,
        older: Annotated[str | None, typer.Option("--older", help="Cursor from a previous call.")] = None,
        full: Annotated[bool, typer.Option("--full", help="Do not cut long messages.")] = False,
    ) -> None:
        """Show the newest messages of a chat."""
        page = api().read(chat, limit=limit, since=parse_since(since) if since else None, older=older)
        typer.echo(render.messages(page, full=full))

    @app.command()
    def search(
        query: Annotated[str, typer.Argument(help="Words to search for.")],
        limit: Annotated[int, typer.Option("--limit", "-n")] = 15,
        offset: Annotated[int, typer.Option("--offset")] = 0,
    ) -> None:
        """Search your Teams messages."""
        typer.echo(render.search(api().search(query, offset=offset, limit=limit)))

    @app.command()
    def files(
        chat: ChatArg,
        name: Annotated[str | None, typer.Option("--name", help="Read this file's text.")] = None,
    ) -> None:
        """List files shared in a chat, or read one with --name."""
        service = api()
        if name:
            typer.echo(render.file_text(service.file_text(chat, name)))
        else:
            typer.echo(render.shared_files(service.files(chat)))

    @app.command()
    def status(who: Annotated[str, typer.Argument(help="Colleague's name or email.")]) -> None:
        """Live Teams presence and card of a colleague."""
        typer.echo(render.presence(api().status(who)))

    @app.command()
    def browser(
        install_sso: Annotated[
            bool,
            typer.Option(
                "--install-sso",
                help="Open the Web Store page of the Microsoft Single Sign On extension "
                "in the Teams browser, so you can add it with one click.",
            ),
        ] = False,
    ) -> None:
        """Start the Teams browser if needed and show its sign-in state (no token values)."""
        if install_sso:
            get_ctx().teams.open_sso_install()
            typer.echo(
                "The Web Store page is open in the Teams browser. "
                "Click 'Add to Chrome', then sign in to Teams."
            )
            return
        typer.echo(json.dumps(get_ctx().teams.status(), indent=1, ensure_ascii=False))


def _write_commands(app: typer.Typer, api: Callable[[], TeamsApi]) -> None:
    """Commands that change things: write, send, mark-read, react, edit, delete."""

    @app.command()
    def write(
        body: Annotated[str, typer.Argument(help="Message (Markdown); '-' reads stdin; '@[Name]' mentions.")],
        chat: Annotated[str | None, typer.Option("--chat", "-c", help="Chat ref or name.")] = None,
        to: Annotated[str | None, typer.Option("--to", "-t", help="Person: name or email.")] = None,
        reply_to: Annotated[
            str | None, typer.Option("--reply-to", "-r", help="Message ref to reply to.")
        ] = None,
        body_format: Annotated[BodyFormat, typer.Option("--format", "-f")] = "markdown",
        at: At = None,
        yes: Yes = False,
        draft: KeepDraft = False,
    ) -> None:
        """Write a message or reply (preview, then confirm)."""
        service = api()
        saved = service.draft(read_body(body), chat=chat, to=to, reply_to=reply_to, body_format=body_format)
        confirm_send(
            render.draft(saved),
            lambda: service.send(saved.ref, send_at=parse_when(at) if at else None),
            yes=yes,
            keep_draft=draft,
            send_command=f"m365 teams send {saved.ref}",
        )

    @app.command()
    def send(draft_ref: Annotated[str, typer.Argument(help="Draft ref (d:…).")], at: At = None) -> None:
        """Send a saved draft now (or schedule it with --at)."""
        typer.echo(api().send(draft_ref, send_at=parse_when(at) if at else None))

    @app.command("mark-read")
    def mark_read(chat: ChatArg) -> None:
        """Mark a chat read up to its newest message."""
        typer.echo(api().mark_read(chat))

    @app.command()
    def react(
        message_ref: Annotated[str, typer.Argument(help="Message ref (p:…).")],
        emoji: Annotated[str, typer.Argument(help="like, heart, laugh, surprised, sad, angry.")],
        remove: Annotated[bool, typer.Option("--remove", help="Remove the reaction.")] = False,
        yes: Yes = False,
    ) -> None:
        """React to a message (or remove the reaction)."""
        if not yes and not typer.confirm(f"{'Remove' if remove else 'Add'} reaction {emoji}?", default=False):
            typer.echo("Cancelled.")
            return
        typer.echo(api().update_message(message_ref, react=emoji, remove_react=remove))

    @app.command()
    def edit(
        message_ref: Annotated[str, typer.Argument(help="Message ref (p:…) of your own message.")],
        body: Annotated[str, typer.Argument(help="New text (Markdown); '-' reads stdin.")],
        yes: Yes = False,
    ) -> None:
        """Edit one of your own messages."""
        if not yes and not typer.confirm("Edit this message?", default=False):
            typer.echo("Cancelled.")
            return
        typer.echo(api().update_message(message_ref, edit=read_body(body)))

    @app.command()
    def delete(
        message_ref: Annotated[str, typer.Argument(help="Message ref (p:…) of your own message.")],
        yes: Yes = False,
    ) -> None:
        """Delete one of your own messages."""
        if not yes and not typer.confirm("Delete this message?", default=False):
            typer.echo("Cancelled.")
            return
        typer.echo(api().update_message(message_ref, delete=True))
