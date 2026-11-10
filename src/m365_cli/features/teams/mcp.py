from __future__ import annotations

from typing import Annotated

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from m365_cli.core.feature import ContextProvider
from m365_cli.core.markup import BodyFormat
from m365_cli.core.mcp_support import LOCAL_WRITE, READ_ONLY, SEND_META, SENDS, ResponseFormat, tool_errors
from m365_cli.core.timeutil import WHEN_FORMATS, parse_since, parse_when
from m365_cli.features.teams import render
from m365_cli.features.teams.api import TeamsApi, teams_api
from m365_cli.features.teams.models import ChatKind

INSTRUCTIONS = (
    "Teams: chat refs look like c:7KQ2PXA, message refs p:…, draft refs d:…. Tools accept a chat as a ref or "
    "by name (chat title, or the person of a 1:1 chat). Sending is two-step: teams_draft returns a "
    "preview; show it to the user and call teams_send only after they explicitly approve."
)

ChatArg = Annotated[
    str,
    Field(description="Chat ref (c:…), or a name: the chat title, or the person of a 1:1 chat."),
]


def register_mcp(server: MCPServer, get_ctx: ContextProvider) -> None:
    """Register the Teams tools."""

    def api() -> TeamsApi:
        return teams_api(get_ctx())

    @server.tool(title="Find Teams chats", annotations=READ_ONLY, structured_output=False)
    @tool_errors
    def teams_find(
        query: Annotated[
            str | None, Field(description="Words in the chat title or the person's name of a 1:1 chat.")
        ] = None,
        kind: Annotated[ChatKind | None, Field(description="chat (1:1) | group | meeting | channel.")] = None,
        unread_only: Annotated[bool, Field(description="Only chats with unread messages.")] = False,
        offset: Annotated[int, Field(ge=0, description="Skip this many chats (paging).")] = 0,
        limit: Annotated[int, Field(ge=1, le=50, description="Number of chats.")] = 15,
    ) -> str:
        """List Teams chats, newest activity first: ref, kind, title, time, unread mark, last message.

        Deleted teams and channels you have left are not listed.
        """
        return render.chats(
            api().chats(query=query, kind=kind, unread_only=unread_only, offset=offset, limit=limit)
        )

    @server.tool(title="Read Teams chat", annotations=READ_ONLY, structured_output=False)
    @tool_errors
    def teams_read(
        chat: ChatArg,
        limit: Annotated[int, Field(ge=1, le=100, description="Newest N messages.")] = 20,
        since: Annotated[
            str | None, Field(description="Only messages after this, e.g. '2d', 'today'.")
        ] = None,
        older: Annotated[
            str | None, Field(description="Cursor from a previous call for the next-older messages.")
        ] = None,
        response_format: Annotated[
            ResponseFormat, Field(description="concise cuts very long messages; detailed shows all.")
        ] = "concise",
    ) -> str:
        """Read the newest messages of a chat, oldest first, with message refs for teams_draft(reply_to=…).

        Shows sender, time, text, the quoted message of a reply, reactions and file names.
        """
        page = api().read(chat, limit=limit, since=parse_since(since) if since else None, older=older)
        return render.messages(page, full=response_format == "detailed")

    @server.tool(title="Teams status", annotations=READ_ONLY, structured_output=False)
    @tool_errors
    def teams_status(
        who: Annotated[str, Field(description="A colleague's name or email.")],
    ) -> str:
        """Live Teams presence of a colleague: available, busy, in a meeting, away, out of office, offline.

        Also shows their status note, out-of-office note and directory card (title, department, office).
        More current than calendar free/busy. An ambiguous name returns an error; use a full name or email.
        """
        return render.presence(api().status(who))

    @server.tool(title="Draft Teams message", annotations=LOCAL_WRITE, structured_output=False)
    @tool_errors
    def teams_draft(
        body: Annotated[
            str,
            Field(
                description="Message text (Markdown by default). '@[Full Name]' mentions that chat member."
            ),
        ],
        chat: Annotated[str | None, Field(description="Chat ref or name (see teams_find).")] = None,
        to: Annotated[
            str | None,
            Field(
                description="Instead of chat: a person's name (their 1:1 chat) or email (also a new chat)."
            ),
        ] = None,
        reply_to: Annotated[
            str | None, Field(description="Message ref (p:…) to reply to; the reply quotes it.")
        ] = None,
        body_format: Annotated[BodyFormat, Field(description="markdown | text | html.")] = "markdown",
    ) -> str:
        """Prepare a Teams message and show its preview. Nothing is sent; the draft stays on this computer.

        Give chat, to, or only reply_to. Then show the preview to the user and call teams_send after
        their approval.
        """
        draft = api().draft(body, chat=chat, to=to, reply_to=reply_to, body_format=body_format)
        return render.draft(draft)

    @server.tool(title="Send Teams message", annotations=SENDS, meta=SEND_META, structured_output=False)
    @tool_errors
    def teams_send(
        draft_ref: Annotated[str, Field(description="Draft ref from teams_draft, e.g. 'd:QX3B7RA'.")],
        send_at: Annotated[
            str | None, Field(description=f"Schedule instead of sending now, e.g. {WHEN_FORMATS}.")
        ] = None,
    ) -> str:
        """Send (or schedule) a draft from teams_draft. Only after the user explicitly approved its preview.

        The user is asked to approve every call; a per-minute send limit applies.
        """
        return api().send(draft_ref, send_at=parse_when(send_at) if send_at else None)

    @server.tool(title="Mark Teams chat read", annotations=LOCAL_WRITE, structured_output=False)
    @tool_errors
    def teams_mark_read(chat: ChatArg) -> str:
        """Mark a chat or channel read up to its newest message. Only your read state changes."""
        return api().mark_read(chat)

    @server.tool(title="Update Teams message", annotations=SENDS, meta=SEND_META, structured_output=False)
    @tool_errors
    def teams_update(
        message_ref: Annotated[str, Field(description="Message ref (p:…) from teams_read.")],
        react: Annotated[
            str | None,
            Field(description="Add a reaction: like, heart, laugh, surprised, sad, angry."),
        ] = None,
        remove_react: Annotated[
            bool, Field(description="Remove that reaction instead of adding it.")
        ] = False,
        edit: Annotated[str | None, Field(description="New text for your own message (Markdown).")] = None,
        delete: Annotated[bool, Field(description="Delete your own message.")] = False,
    ) -> str:
        """React to a message, or edit/delete one of your own. Visible to others, so it needs approval.

        Give exactly one of react, edit or delete. The user is asked before each call.
        """
        return api().update_message(
            message_ref, react=react, remove_react=remove_react, edit=edit, delete=delete
        )
