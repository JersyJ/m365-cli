from __future__ import annotations

from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field


class _Model(BaseModel):
    model_config = ConfigDict(frozen=True)


class Person(_Model):
    """Address-book entry."""

    name: str
    email: str = ""
    title: str = ""
    department: str = ""
    office: str = ""
    phone: str = ""
    mobile: str = ""
    manager: str = ""


class Availability(_Model):
    """Availability of a person.

    Attributes:
        now: free | busy | tentative | out of office | working elsewhere | unknown.
        current_until: End of the current state (None when free).
        out_of_office: Out-of-office periods from now on.
        busy_today: Busy/tentative periods for the rest of today.
    """

    now: str
    current_until: datetime | None = None
    out_of_office: list[tuple[datetime, datetime]] = Field(default_factory=list)
    busy_today: list[tuple[datetime, datetime]] = Field(default_factory=list)
