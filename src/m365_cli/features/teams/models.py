from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

ChatKind = Literal["chat", "group", "meeting", "channel"]
CHAT_KINDS: tuple[ChatKind, ...] = ("chat", "group", "meeting", "channel")


class Member(BaseModel):
    """A person in a chat: display name and Teams user ID (MRI, e.g. '8:orgid:<guid>')."""

    model_config = ConfigDict(frozen=True)

    name: str
    mri: str
    email: str = ""


class Chat(BaseModel):
    """A Teams conversation: 1:1 chat, group chat, meeting chat or channel."""

    model_config = ConfigDict(frozen=True)

    ref: str
    id: str
    kind: ChatKind
    title: str
    last_time: datetime | None = None
    last_from: str = ""
    last_preview: str = ""
    unread: bool = False


class TeamsMessage(BaseModel):
    """One message, already converted to plain text.

    Attributes:
        quote: "Sender: preview" of the message this one replies to ('' when none).
        reactions: Reaction key (e.g. 'like', '1f525_fire') to number of people.
        system: An event like "X added Y", not a message somebody wrote.
    """

    model_config = ConfigDict(frozen=True)

    ref: str
    id: str
    sender: str
    sender_mri: str = ""
    time: datetime | None = None
    text: str = ""
    edited: bool = False
    deleted: bool = False
    system: bool = False
    quote: str = ""
    mentions: list[str] = Field(default_factory=list)
    reactions: dict[str, int] = Field(default_factory=dict)
    files: list[str] = Field(default_factory=list)
    images: int = 0


class MessagePage(BaseModel):
    """Messages of one chat, oldest first; `older` is the cursor ref for the previous page."""

    model_config = ConfigDict(frozen=True)

    chat: Chat
    messages: list[TeamsMessage]
    older: str | None = None


class TeamsDraft(BaseModel):
    """A message saved locally until it is sent; Teams itself never sees it before teams_send.

    For a first message to a person without a 1:1 chat, `chat_id` is empty and `to` holds the person.
    """

    ref: str = ""
    id: str
    chat_id: str = ""
    chat_title: str
    chat_kind: ChatKind = "chat"
    to: Member | None = None
    html: str
    text: str
    reply_to_id: str = ""
    reply_quote: str = ""
    reply_to_mri: str = ""
    reply_to_name: str = ""
    mentions: list[Member] = Field(default_factory=list)
    created: datetime


class Presence(BaseModel):
    """A colleague's live Teams presence plus their directory card."""

    model_config = ConfigDict(frozen=True)

    name: str
    mri: str
    email: str = ""
    title: str = ""
    department: str = ""
    office: str = ""
    status: str
    available: bool = False
    out_of_office: bool = False
    ooo_note: str = ""
    message: str = ""
    last_active: datetime | None = None
    device: str = ""
