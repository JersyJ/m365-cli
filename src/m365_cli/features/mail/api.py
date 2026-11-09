from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Protocol

from m365_cli.core.markup import BodyFormat
from m365_cli.core.paging import Page
from m365_cli.features.mail.models import (
    AttachmentText,
    ComposeMode,
    DigestGroup,
    Draft,
    FlagAction,
    MailMessage,
    MailQuery,
    MailSummary,
    WaitingItem,
)

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

KIND = "m"


class MailApi(Protocol):
    """Mail operations the CLI and MCP tools use, independent of the backend.

    Implemented by `MailService` (classic Outlook) and `FakeMailService` (fake backend).
    """

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first. Folder "all" searches every mail folder."""
        ...

    def select(self, query: MailQuery, *, max_items: int) -> list[str]:
        """Refs of up to `max_items` mails matching the query (cheap; for bulk updates)."""
        ...

    def digest(self, query: MailQuery, *, max_items: int = 300) -> list[DigestGroup]:
        """Mails matching the query, grouped by sender (at most `max_items` mails are counted)."""
        ...

    def mailboxes(self) -> list[tuple[str, str]]:
        """Mailboxes available as (name, kind): primary | account | shared | other."""
        ...

    def thread(self, ref: str, *, full: bool = False) -> list[MailMessage]:
        """The whole conversation of a mail across folders, oldest first."""
        ...

    def waiting(
        self, *, days: int = 3, within_days: int = 14, limit: int = 20, mailbox: str | None = None
    ) -> list[WaitingItem]:
        """Sent mails from the last `within_days` days, older than `days`, with no reply yet."""
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
        body_format: BodyFormat = "markdown",
        attachments: list[str] | None = None,
        from_mailbox: str | None = None,
    ) -> Draft:
        """Save a draft (in Drafts) and return its preview. Never sends."""
        ...

    def send(self, ref: str, *, send_at: datetime | None = None) -> str:
        """Send a draft now or, with `send_at`, schedule it (rate-limited and logged)."""
        ...

    def update(
        self,
        refs: list[str],
        *,
        read: bool | None = None,
        flag: FlagAction | None = None,
        categories: list[str] | None = None,
        move_to: str | None = None,
        delete: bool = False,
    ) -> str:
        """Mark read/unread, flag, set categories, move or delete one or more mails; returns a summary."""
        ...

    def attachments(self, ref: str, *, save_to: str | None = None, include_inline: bool = False) -> list[str]:
        """List the attachments, or save them into directory `save_to`."""
        ...

    def attachment_text(self, ref: str, which: str) -> AttachmentText:
        """Extracted text of one attachment, chosen by 1-based index or by file name."""
        ...


def mail_api(ctx: AppContext) -> MailApi:
    """Mail operations of the active backend."""
    if ctx.is_fake:
        from m365_cli.features.mail.fake import FakeMailService

        return FakeMailService(ctx)
    from m365_cli.features.mail.service import MailService

    return MailService(ctx)
