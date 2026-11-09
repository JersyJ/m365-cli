from __future__ import annotations

from datetime import datetime
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

ComposeMode = Literal["new", "reply", "reply_all", "forward"]
FlagAction = Literal["set", "clear", "complete"]


class FlagChoice(StrEnum):
    """CLI-facing mirror of `FlagAction`."""

    SET = "set"
    CLEAR = "clear"
    COMPLETE = "complete"


class _Model(BaseModel):
    model_config = ConfigDict(frozen=True)


class MailQuery(_Model):
    """Mail search; all given conditions must match. `text` matches subject or body."""

    folder: str = "inbox"
    unread_only: bool = False
    since: datetime | None = None
    until: datetime | None = None
    sender: str | None = None
    subject: str | None = None
    text: str | None = None
    has_attachments: bool | None = None
    flagged: bool | None = None
    limit: int = Field(default=15, ge=1, le=200)


class Attachment(_Model):
    """Attachment; `index` is 1-based, as in Outlook."""

    index: int
    name: str
    size: int
    inline: bool = False


class MailSummary(_Model):
    """Mail in a list. `kind` is mail | meeting | report | other; `preview` is set on request."""

    ref: str
    received: datetime | None
    sender: str
    subject: str
    unread: bool = False
    has_attachments: bool = False
    flagged: bool = False
    kind: str = "mail"
    preview: str = ""


class MailMessage(MailSummary):
    """Mail with its body; `quoted_cut`/`truncated` say what was left out."""

    sender_email: str = ""
    to: str = ""
    cc: str = ""
    categories: str = ""
    attachments: list[Attachment] = Field(default_factory=list)
    body: str = ""
    quoted_cut: bool = False
    truncated: bool = False


class Draft(_Model):
    """Saved, unsent draft. `text` is what we wrote, without the quoted original."""

    ref: str
    mode: ComposeMode
    to: list[str]
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    subject: str
    text: str
    includes_original: bool = False
    attachments: list[str] = Field(default_factory=list)
