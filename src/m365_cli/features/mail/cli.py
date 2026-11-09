from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, cast

import typer

from m365_cli.core.cli_support import confirm_send, read_body, split_list
from m365_cli.core.feature import ContextProvider
from m365_cli.core.markup import BodyFormat
from m365_cli.core.timeutil import parse_since, parse_when
from m365_cli.features.mail import render
from m365_cli.features.mail.api import MailApi, mail_api
from m365_cli.features.mail.models import Draft, FlagAction, FlagChoice, MailQuery

Ref = Annotated[str, typer.Argument(help="Mail ref, e.g. m:7KQ2PXA (from `mail list`).")]
Body = Annotated[str, typer.Option("--body", "-b", help="Text; '-' reads stdin.")]
To = Annotated[list[str] | None, typer.Option("--to", "-t", help="Recipient(s); repeat or comma-separate.")]
Cc = Annotated[list[str] | None, typer.Option("--cc", help="Cc recipient(s).")]
Bcc = Annotated[list[str] | None, typer.Option("--bcc", help="Bcc recipient(s).")]
Attach = Annotated[list[str] | None, typer.Option("--attach", "-a", help="File to attach (repeatable).")]
Fmt = Annotated[BodyFormat, typer.Option("--format", help="markdown (default) | html | text.")]
Yes = Annotated[bool, typer.Option("--yes", "-y", help="Send without asking.")]
Box = Annotated[str | None, typer.Option("--mailbox", "-m", help="Shared mailbox or other account.")]
FromBox = Annotated[
    str | None, typer.Option("--from-mailbox", help="Send from this account / on behalf of a shared mailbox.")
]
KeepDraft = Annotated[bool, typer.Option("--draft", help="Only save the draft, do not send.")]


def build_cli(get_ctx: ContextProvider) -> typer.Typer:
    """`m365 mail …` commands."""
    app = typer.Typer(help="Read, search, compose and manage mail.", no_args_is_help=True)

    def api() -> MailApi:
        return mail_api(get_ctx())

    def compose(make: Callable[[MailApi], Draft], *, yes: bool, keep: bool) -> None:
        service = api()
        d = make(service)
        confirm_send(
            render.draft(d),
            lambda: service.send(d.ref),
            yes=yes,
            keep_draft=keep,
            send_command=f"m365 mail send {d.ref}",
        )

    _reading_commands(app, api, get_ctx)
    _writing_commands(app, api, compose)
    return app


def _reading_commands(app: typer.Typer, api: Callable[[], MailApi], get_ctx: ContextProvider) -> None:
    """Commands that only read: list, read, digest, thread, waiting."""

    @app.command("list")
    def list_(
        mailbox: Box = None,
        folder: Annotated[
            str, typer.Option("--folder", "-f", help="inbox|sent|drafts|… or 'Inbox/Sub'.")
        ] = "inbox",
        unread: Annotated[bool, typer.Option("--unread", "-u", help="Only unread.")] = False,
        since: Annotated[
            str | None, typer.Option("--since", "-s", help="'2d', '12h', 'yesterday', date.")
        ] = None,
        until: Annotated[str | None, typer.Option("--until", help="Upper bound (date/time).")] = None,
        sender: Annotated[str | None, typer.Option("--from", help="Sender name/address contains.")] = None,
        subject: Annotated[str | None, typer.Option("--subject", help="Subject contains.")] = None,
        text: Annotated[str | None, typer.Option("--text", "-q", help="Subject or body contains.")] = None,
        attachments: Annotated[bool | None, typer.Option("--attachments/--no-attachments")] = None,
        flagged: Annotated[bool | None, typer.Option("--flagged/--not-flagged")] = None,
        limit: Annotated[int | None, typer.Option("--limit", "-n", help="Mails per page.")] = None,
        offset: Annotated[int, typer.Option("--offset", help="Skip this many mails (next page).")] = 0,
        preview: Annotated[
            bool, typer.Option("--preview", "-p", help="Show the start of each body.")
        ] = False,
    ) -> None:
        """List/search mail, newest first."""
        query = MailQuery(
            folder=folder,
            mailbox=mailbox,
            unread_only=unread,
            since=parse_since(since) if since else None,
            until=parse_when(until) if until else None,
            sender=sender,
            subject=subject,
            text=text,
            has_attachments=attachments,
            flagged=flagged,
            limit=limit or get_ctx().settings.list_limit,
        )
        typer.echo(render.summaries(api().find(query, offset=offset, preview=preview)))

    @app.command()
    def read(
        ref: Ref, full: Annotated[bool, typer.Option("--full", help="Include quoted history.")] = False
    ) -> None:
        """Read one mail as plain text."""
        typer.echo(render.message(api().read(ref, full=full)))

    @app.command()
    def digest(
        mailbox: Box = None,
        folder: Annotated[str, typer.Option("--folder", "-f", help="Folder, path or 'all'.")] = "inbox",
        since: Annotated[str, typer.Option("--since", "-s", help="'1d', '12h', 'mon', date.")] = "1d",
        unread: Annotated[bool, typer.Option("--unread", "-u", help="Only unread.")] = False,
    ) -> None:
        """Overview grouped by sender: counts, unread, newest time, top subjects."""
        query = MailQuery(folder=folder, mailbox=mailbox, unread_only=unread, since=parse_since(since))
        typer.echo(render.digest(api().digest(query)))

    @app.command()
    def thread(
        ref: Ref, full: Annotated[bool, typer.Option("--full", help="Keep quoted history.")] = False
    ) -> None:
        """Show the whole conversation of a mail, oldest first."""
        typer.echo(render.thread(api().thread(ref, full=full)))

    @app.command()
    def waiting(
        mailbox: Box = None,
        days: Annotated[int, typer.Option("--days", "-d", help="Sent at least this many days ago.")] = 3,
        within: Annotated[int, typer.Option("--within", help="Look back this many days.")] = 14,
    ) -> None:
        """List sent mails that got no reply yet."""
        typer.echo(render.waiting(api().waiting(days=days, within_days=within, mailbox=mailbox)))

    @app.command()
    def mailboxes() -> None:
        """List the mailboxes you can use with --mailbox / --from-mailbox."""
        typer.echo("\n".join(f"{name} ({kind})" for name, kind in api().mailboxes()))


def _writing_commands(app: typer.Typer, api: Callable[[], MailApi], compose: Callable[..., None]) -> None:
    """Commands that change or send mail: new, reply, forward, send, update, attachments."""

    @app.command()
    def new(
        to: To,
        subject: Annotated[str, typer.Option("--subject", "-s")],
        body: Body,
        cc: Cc = None,
        bcc: Bcc = None,
        attach: Attach = None,
        body_format: Fmt = "markdown",
        from_mailbox: FromBox = None,
        yes: Yes = False,
        draft: KeepDraft = False,
    ) -> None:
        """Write a new mail (preview, then confirm)."""
        compose(
            lambda s: s.compose(
                "new",
                to=split_list(to),
                cc=split_list(cc),
                bcc=split_list(bcc),
                subject=subject,
                body=read_body(body),
                attachments=attach,
                from_mailbox=from_mailbox,
                body_format=body_format,
            ),
            yes=yes,
            keep=draft,
        )

    @app.command()
    def reply(
        ref: Ref,
        body: Body,
        all_: Annotated[bool, typer.Option("--all", help="Reply to all.")] = False,
        to: To = None,
        cc: Cc = None,
        attach: Attach = None,
        body_format: Fmt = "markdown",
        from_mailbox: FromBox = None,
        yes: Yes = False,
        draft: KeepDraft = False,
    ) -> None:
        """Reply (or --all) to a mail; extra --to/--cc are added."""
        compose(
            lambda s: s.compose(
                "reply_all" if all_ else "reply",
                ref=ref,
                to=split_list(to),
                cc=split_list(cc),
                body=read_body(body),
                attachments=attach,
                from_mailbox=from_mailbox,
                body_format=body_format,
            ),
            yes=yes,
            keep=draft,
        )

    @app.command()
    def forward(
        ref: Ref,
        to: To,
        body: Body,
        cc: Cc = None,
        attach: Attach = None,
        body_format: Fmt = "markdown",
        from_mailbox: FromBox = None,
        yes: Yes = False,
        draft: KeepDraft = False,
    ) -> None:
        """Forward a mail."""
        compose(
            lambda s: s.compose(
                "forward",
                ref=ref,
                to=split_list(to),
                cc=split_list(cc),
                body=read_body(body),
                attachments=attach,
                from_mailbox=from_mailbox,
                body_format=body_format,
            ),
            yes=yes,
            keep=draft,
        )

    @app.command()
    def send(
        ref: Ref,
        at: Annotated[str | None, typer.Option("--at", help="Schedule, e.g. 'tomorrow 8:00'.")] = None,
        yes: Yes = False,
    ) -> None:
        """Send (or schedule with --at) an existing draft."""
        send_at = parse_when(at) if at else None
        question = f"Schedule draft {ref} for {at}?" if at else f"Send draft {ref}?"
        if yes or typer.confirm(question, default=False):
            typer.echo(api().send(ref, send_at=send_at))

    @app.command()
    def update(
        mailbox: Box = None,
        refs: Annotated[list[str] | None, typer.Argument(help="Mail refs (or use filter options).")] = None,
        folder: Annotated[
            str | None, typer.Option("--folder", "-f", help="Filter: folder, path or 'all'.")
        ] = None,
        sender: Annotated[str | None, typer.Option("--from", help="Filter: sender contains.")] = None,
        subject: Annotated[str | None, typer.Option("--subject", help="Filter: subject contains.")] = None,
        unread_only: Annotated[bool, typer.Option("--unread-only", help="Filter: only unread.")] = False,
        since: Annotated[str | None, typer.Option("--since", help="Filter: received at/after.")] = None,
        max_items: Annotated[int, typer.Option("--max", help="Upper limit for filter updates.")] = 50,
        read: Annotated[bool | None, typer.Option("--read/--unread")] = None,
        flag: Annotated[FlagChoice | None, typer.Option("--flag")] = None,
        category: Annotated[
            list[str] | None, typer.Option("--category", "-c", help="Replace categories.")
        ] = None,
        move: Annotated[str | None, typer.Option("--move", help="Target folder.")] = None,
        delete: Annotated[bool, typer.Option("--delete", help="Move to Deleted Items.")] = False,
        yes: Yes = False,
    ) -> None:
        """Mark read/unread, flag, categorize, move or delete mails given by ref or by filter."""
        service = api()
        targets = list(refs or [])
        if not targets:
            query = MailQuery(
                folder=folder or "inbox",
                mailbox=mailbox,
                sender=sender,
                subject=subject,
                unread_only=unread_only,
                since=parse_since(since) if since else None,
                limit=min(max_items, 200),
            )
            targets = service.select(query, max_items=max_items)
            if not targets or not (yes or typer.confirm(f"Update {len(targets)} mail(s)?", default=False)):
                typer.echo("Nothing changed.")
                return
        typer.echo(
            service.update(
                targets,
                read=read,
                flag=cast(FlagAction | None, flag.value if flag else None),
                categories=split_list(category) if category else None,
                move_to=move,
                delete=delete,
            )
        )

    @app.command()
    def attachments(
        ref: Ref,
        read: Annotated[
            str | None, typer.Option("--read", help="Show the text of this attachment (number or name).")
        ] = None,
        save: Annotated[str | None, typer.Option("--save", help="Directory to save into.")] = None,
        inline: Annotated[bool, typer.Option("--inline", help="Include inline images.")] = False,
    ) -> None:
        """List, read or save attachments."""
        if read is not None:
            typer.echo(render.attachment_text(api().attachment_text(ref, read)))
            return
        typer.echo(
            "\n".join(api().attachments(ref, save_to=save, include_inline=inline)) or "(no attachments)"
        )
