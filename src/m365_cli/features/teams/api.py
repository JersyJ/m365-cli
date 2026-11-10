from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Protocol

from m365_cli.core.markup import BodyFormat
from m365_cli.core.paging import Page
from m365_cli.features.teams.models import Chat, ChatKind, Digest, MessagePage, Presence, TeamsDraft

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext


class TeamsApi(Protocol):
    """Teams operations the CLI and MCP tools use, independent of the backend.

    A chat is given as a ref ('c:…'), a Teams conversation ID ('19:…'), or a name: the chat title or,
    for 1:1 chats, the other person's name.
    """

    def chats(
        self,
        *,
        query: str | None = None,
        kind: ChatKind | None = None,
        unread_only: bool = False,
        offset: int = 0,
        limit: int = 15,
    ) -> Page[Chat]:
        """Chats, newest activity first; `query` matches the title (or the person of a 1:1 chat)."""
        ...

    def read(
        self, chat: str, *, limit: int = 20, since: datetime | None = None, older: str | None = None
    ) -> MessagePage:
        """Newest `limit` messages of a chat, oldest first; `older` is a cursor ref from a previous page.

        Raises:
            InvalidInputError: The chat name matches no chat or several chats.
        """
        ...

    def draft(
        self,
        body: str,
        *,
        chat: str | None = None,
        to: str | None = None,
        reply_to: str | None = None,
        body_format: BodyFormat = "markdown",
    ) -> TeamsDraft:
        """Save a message locally (not sent) and return it for the preview.

        Give `chat`, or `to` (a person: their 1:1 chat, created on send when there is none yet), or
        only `reply_to` (a message ref; the reply goes to its chat and quotes it). `@[Name]` in the
        body mentions that chat member.
        """
        ...

    def send(self, draft_ref: str, *, send_at: datetime | None = None) -> str:
        """Send a draft now, or schedule it with `send_at`; returns a short confirmation.

        Raises:
            RateLimitError: The per-minute send limit is reached.
            InvalidInputError: The chat is protected, or `send_at` is in the past.
        """
        ...

    def status(self, query: str) -> Presence:
        """Live Teams presence and directory card of a colleague (by name or email).

        Raises:
            InvalidInputError: The name matches no one or several people.
        """
        ...

    def mark_read(self, chat: str) -> str:
        """Mark a chat read up to its newest message; returns a short confirmation."""
        ...

    def digest(self, *, unread_only: bool = True, limit: int = 15) -> Digest:
        """What needs you on Teams: @mentions/replies to you, plus unread conversations."""
        ...

    def update_message(
        self,
        message_ref: str,
        *,
        react: str | None = None,
        remove_react: bool = False,
        edit: str | None = None,
        delete: bool = False,
        body_format: BodyFormat = "markdown",
    ) -> str:
        """React to, edit, or delete a message. Edit and delete work only on your own messages.

        Raises:
            InvalidInputError: No action given, or more than one kind at once.
        """
        ...


def teams_api(ctx: AppContext) -> TeamsApi:
    """Teams operations of the active backend."""
    if ctx.is_fake:
        from m365_cli.features.teams.fake import FakeTeamsService

        return FakeTeamsService(ctx)
    from m365_cli.features.teams.service import TeamsService

    return TeamsService(ctx)
