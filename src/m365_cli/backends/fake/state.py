from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field

FlagState = Literal["none", "set", "complete"]


class FakeAttachment(BaseModel):
    """Attachment of a fake mail."""

    name: str
    size: int
    inline: bool = False
    text: str = ""


class FakeMail(BaseModel):
    """Fake mail, draft or meeting request."""

    id: str
    folder: str = "inbox"
    received: datetime
    sender_name: str
    sender_email: str
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    subject: str
    body: str = ""
    unread: bool = False
    flag: FlagState = "none"
    categories: list[str] = Field(default_factory=list)
    attachments: list[FakeAttachment] = Field(default_factory=list)
    kind: str = "mail"
    is_draft: bool = False
    draft_mode: str | None = None
    source_id: str | None = None
    event_id: str | None = None
    protected: bool = False
    mailbox: str = ""
    from_mailbox: str = ""


class FakeBlock(BaseModel):
    """Free/busy block; `status` uses the Outlook values (1 tentative, 2 busy, 3 out of office)."""

    start: datetime
    end: datetime
    status: int


class FakeEvent(BaseModel):
    """Fake calendar item."""

    id: str
    start: datetime
    end: datetime
    all_day: bool = False
    subject: str
    location: str = ""
    organizer: str = ""
    organizer_is_me: bool = False
    required: list[str] = Field(default_factory=list)
    optional: list[str] = Field(default_factory=list)
    rooms: list[str] = Field(default_factory=list)
    body: str = ""
    show_as: int = 2
    my_response: str = ""
    recurring: bool = False
    canceled: bool = False
    is_meeting: bool = True
    sent: bool = True
    join_url: str = ""
    repeat: dict[str, object] | None = None


class FakePerson(BaseModel):
    """Fake address-book entry.

    `blocks` is None when no free/busy is available (e.g. external); `calendar` holds the events of a
    calendar shared with the user (None = not shared).
    """

    name: str
    email: str
    title: str = ""
    department: str = ""
    office: str = ""
    phone: str = ""
    mobile: str = ""
    manager: str = ""
    blocks: list[FakeBlock] | None = Field(default_factory=list)
    calendar: list[FakeEvent] | None = None


class FakeTeamsMessage(BaseModel):
    """Fake Teams message; `html` is the message content as Teams stores it."""

    id: str
    sender: str
    time: datetime
    html: str
    edited: bool = False
    deleted: bool = False
    system: bool = False
    reactions: dict[str, int] = Field(default_factory=dict)
    files: list[str] = Field(default_factory=list)


class FakeChat(BaseModel):
    """Fake Teams chat; `members` are names of people in `FakeState.people` (or the user)."""

    id: str
    kind: Literal["chat", "group", "meeting", "channel"]
    title: str
    members: list[str] = Field(default_factory=list)
    unread: bool = False
    deleted: bool = False
    messages: list[FakeTeamsMessage] = Field(default_factory=list)


class Outgoing(BaseModel):
    """A send action recorded by the fake backend (it never really sends)."""

    at: datetime
    action: str
    item_id: str
    summary: str
    recipients: list[str] = Field(default_factory=list)
    send_at: datetime | None = None


class SavedFile(BaseModel):
    """An attachment the fake backend pretended to save."""

    mail_id: str
    name: str
    path: str


class FakeState(BaseModel):
    """Complete state of the fake backend: mailbox, calendar, address book and what the agent did.

    Tests and evaluations use it instead of Outlook; an evaluation reads the saved file afterwards to
    check the outcome (drafts, flags, outbox). `clock` is frozen, so relative times like "today"
    refer to the day of the data, and a snapshot of real data stays usable for weeks.
    """

    source: Literal["fixtures", "snapshot"] = "fixtures"
    clock: datetime
    me: FakePerson
    folders: list[str]
    mails: list[FakeMail]
    events: list[FakeEvent]
    people: list[FakePerson]
    signatures: dict[str, str] = Field(default_factory=dict)
    mailboxes: list[str] = Field(default_factory=list)
    outbox: list[Outgoing] = Field(default_factory=list)
    saved_files: list[SavedFile] = Field(default_factory=list)
    teams_chats: list[FakeChat] = Field(default_factory=list)
    teams_drafts: dict[str, dict[str, object]] = Field(default_factory=dict)
    next_id: int = 1000

    def new_id(self, prefix: str) -> str:
        """New unique item ID with the given prefix."""
        self.next_id += 1
        return f"{prefix}{self.next_id}"
