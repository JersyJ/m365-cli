from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, cast

import typer

from m365_cli.core.cli_support import confirm_send, read_body, split_list
from m365_cli.core.feature import ContextProvider
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
Html = Annotated[bool, typer.Option("--html", help="Body is HTML.")]
Yes = Annotated[bool, typer.Option("--yes", "-y", help="Send without asking.")]
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

    @app.command("list")
    def list_(
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
    def new(
        to: To,
        subject: Annotated[str, typer.Option("--subject", "-s")],
        body: Body,
        cc: Cc = None,
        bcc: Bcc = None,
        attach: Attach = None,
        html: Html = False,
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
                html=html,
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
        html: Html = False,
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
                html=html,
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
        html: Html = False,
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
                html=html,
            ),
            yes=yes,
            keep=draft,
        )

    @app.command()
    def send(ref: Ref, yes: Yes = False) -> None:
        """Send an existing draft."""
        if yes or typer.confirm(f"Send draft {ref}?", default=False):
            typer.echo(api().send(ref))

    @app.command()
    def update(
        ref: Ref,
        read: Annotated[bool | None, typer.Option("--read/--unread")] = None,
        flag: Annotated[FlagChoice | None, typer.Option("--flag")] = None,
        category: Annotated[
            list[str] | None, typer.Option("--category", "-c", help="Replace categories.")
        ] = None,
        move: Annotated[str | None, typer.Option("--move", help="Target folder.")] = None,
        delete: Annotated[bool, typer.Option("--delete", help="Move to Deleted Items.")] = False,
    ) -> None:
        """Mark read/unread, flag, categorize, move or delete."""
        typer.echo(
            api().update(
                ref,
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
        save: Annotated[str | None, typer.Option("--save", help="Directory to save into.")] = None,
        inline: Annotated[bool, typer.Option("--inline", help="Include inline images.")] = False,
    ) -> None:
        """List or save attachments."""
        typer.echo(
            "\n".join(api().attachments(ref, save_to=save, include_inline=inline)) or "(no attachments)"
        )

    return app
