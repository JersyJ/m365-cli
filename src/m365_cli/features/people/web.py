from __future__ import annotations

from datetime import timedelta
from typing import Any

from m365_cli.core.context import AppContext
from m365_cli.core.errors import UnresolvedRecipientsError
from m365_cli.core.freebusy import parse_freebusy
from m365_cli.core.timeutil import now
from m365_cli.features.people.models import Availability, Person
from m365_cli.features.people.rules import summarize

_FREEBUSY_MINUTES = 30


class WebPeopleService:
    """People on Outlook on the web: the people search (contacts and directory) and Exchange free/busy.

    The manager is not available through this API, so `Person.manager` stays empty.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook web session of `ctx`."""
        self._web = ctx.outlook_web

    def lookup(self, query: str, *, availability_days: int = 0) -> tuple[Person, Availability | None]:
        """Contact card, plus availability when `availability_days` > 0.

        Raises:
            UnresolvedRecipientsError: No unique match.
        """
        person = self._person(query)
        if availability_days <= 0 or not person.email:
            return person, None
        at = now()
        start = at.replace(hour=0, minute=0, second=0, microsecond=0)
        schedules = self._web.schedules(
            [person.email], start, start + timedelta(days=availability_days), _FREEBUSY_MINUTES
        )
        view = schedules[0][1] if schedules else None
        blocks = parse_freebusy(view, start, _FREEBUSY_MINUTES) if view is not None else None
        return person, summarize(blocks, at)

    def _person(self, query: str) -> Person:
        key = query.strip()
        people = [p for p in self._web.find_people(key) if _addresses(p)]
        if "@" in key:
            match = [p for p in people if key.lower() in (a.lower() for a in _addresses(p))]
            return _card(match[0]) if match else Person(name=key, email=key)
        exact = [p for p in people if str(p.get("DisplayName", "")).lower() == key.lower()]
        chosen = exact or people
        if len(chosen) != 1:
            raise UnresolvedRecipientsError([query])
        return _card(chosen[0])


def _addresses(person: dict[str, Any]) -> list[str]:
    return [str(e["Address"]) for e in person.get("EmailAddresses") or [] if e.get("Address")]


def _phone(person: dict[str, Any], kind: str) -> str:
    return next(
        (
            str(p.get("Number") or "")
            for p in person.get("Phones") or []
            if str(p.get("Type", "")).lower() == kind
        ),
        "",
    )


def _card(person: dict[str, Any]) -> Person:
    return Person(
        name=str(person.get("DisplayName") or ""),
        email=_addresses(person)[0],
        title=str(person.get("Title") or person.get("JobTitle") or ""),
        department=str(person.get("Department") or ""),
        office=str(person.get("OfficeLocation") or ""),
        phone=_phone(person, "business"),
        mobile=_phone(person, "mobile"),
    )
