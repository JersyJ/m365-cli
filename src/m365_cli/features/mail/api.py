from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

from m365_cli.core.paging import Page
from m365_cli.features.mail.models import (
    ComposeMode,
    Draft,
    FlagAction,
    MailMessage,
    MailQuery,
    MailSummary,
)

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

KIND = "m"


class MailApi(Protocol):
    """Mail operations the CLI and MCP tools use, independent of the backend.

    Implemented by `MailService` (classic Outlook) and `FakeMailService` (fake backend).
    """

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first."""
        ...

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        ...

    def compose(
        self,
        mode: ComposeMode,
        *,
        body: str,
        ref: str | None = None,
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        subject: str | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
    ) -> Draft:
        """Save a draft (in Drafts) and return its preview. Never sends."""
        ...

    def send(self, ref: str) -> str:
        """Send a draft (rate-limited and logged)."""
        ...

    def update(
        self,
        ref: str,
        *,
        read: bool | None = None,
        flag: FlagAction | None = None,
        categories: list[str] | None = None,
        move_to: str | None = None,
        delete: bool = False,
    ) -> str:
        """Mark read/unread, flag, set categories, move or delete; returns a short result."""
        ...

    def attachments(self, ref: str, *, save_to: str | None = None, include_inline: bool = False) -> list[str]:
        """List the attachments, or save them into directory `save_to`."""
        ...


def mail_api(ctx: AppContext) -> MailApi:
    """Mail operations of the active backend."""
    if ctx.is_fake:
        from m365_cli.features.mail.fake import FakeMailService

        return FakeMailService(ctx)
    from m365_cli.features.mail.service import MailService

    return MailService(ctx)
