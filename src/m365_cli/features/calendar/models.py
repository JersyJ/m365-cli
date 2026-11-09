from __future__ import annotations

from datetime import datetime
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

Response = Literal["accept", "tentative", "decline"]


class ResponseChoice(StrEnum):
    """CLI-facing mirror of `Response`."""

    ACCEPT = "accept"
    TENTATIVE = "tentative"
    DECLINE = "decline"


class _Model(BaseModel):
    model_config = ConfigDict(frozen=True)


class EventSummary(_Model):
    """Calendar item in a list.

    Attributes:
        show_as: free | tentative | busy | out of office | working elsewhere.
        my_response: organizer | accepted | tentative | declined | not responded | "" (no meeting).
    """

    ref: str
    start: datetime
    end: datetime
    all_day: bool = False
    subject: str
    location: str = ""
    show_as: str = "busy"
    organizer: str = ""
    my_response: str = ""
    recurring: bool = False
    canceled: bool = False
    join_url: str = ""


class Event(EventSummary):
    """Calendar item with attendees and description."""

    required: str = ""
    optional: str = ""
    categories: str = ""
    body: str = ""
    truncated: bool = False


class MeetingDraft(_Model):
    """Saved calendar item.

    Attributes:
        is_meeting: True = meeting, saved but not sent yet. False = personal appointment (already
            saved, nothing to send).
    """

    ref: str
    subject: str
    start: datetime
    end: datetime
    location: str = ""
    required: list[str] = Field(default_factory=list)
    optional: list[str] = Field(default_factory=list)
    rooms: list[str] = Field(default_factory=list)
    text: str = ""
    is_meeting: bool = True


class Slot(_Model):
    """Free period."""

    start: datetime
    end: datetime


class FreeTimeResult(_Model):
    """Free periods, plus notes about attendees without free/busy data."""

    slots: list[Slot]
    notes: list[str] = Field(default_factory=list)
