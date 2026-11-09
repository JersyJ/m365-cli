from __future__ import annotations

from datetime import timedelta

from m365_cli.backends.fake.directory import resolve_person
from m365_cli.core.context import AppContext
from m365_cli.core.freebusy import Block, BusyStatus
from m365_cli.core.timeutil import now
from m365_cli.features.people.models import Availability, Person
from m365_cli.features.people.rules import summarize


class FakePeopleService:
    """People operations on the fake backend."""

    def __init__(self, ctx: AppContext) -> None:
        """Use the fake store of `ctx`."""
        self._store = ctx.fake

    def lookup(self, query: str, *, availability_days: int = 0) -> tuple[Person, Availability | None]:
        """Contact card, plus availability when `availability_days` > 0."""
        p = resolve_person(self._store.state, query)
        person = Person(
            name=p.name,
            email=p.email,
            title=p.title,
            department=p.department,
            office=p.office,
            phone=p.phone,
            mobile=p.mobile,
            manager=p.manager,
        )
        if availability_days <= 0:
            return person, None
        if p.blocks is None:
            return person, summarize(None, now())
        at = now()
        end = at + timedelta(days=availability_days)
        blocks = [Block(b.start, b.end, BusyStatus(b.status)) for b in p.blocks if b.start < end]
        return person, summarize(blocks, at)
