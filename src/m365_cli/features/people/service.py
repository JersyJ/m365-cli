from __future__ import annotations

import contextlib

import pywintypes

from m365_cli.backends.outlook import Com, Outlook
from m365_cli.backends.outlook.freebusy import recipient_blocks
from m365_cli.backends.outlook.items import get
from m365_cli.core.context import AppContext
from m365_cli.core.timeutil import now
from m365_cli.features.people.models import Availability, Person
from m365_cli.features.people.rules import summarize

_FREEBUSY_MINUTES = 30


def _person(ol: Outlook, recipient: Com) -> Person:
    entry = recipient.AddressEntry
    user = None
    with contextlib.suppress(pywintypes.com_error):
        user = entry.GetExchangeUser()
    if user is None:
        return Person(name=recipient.Name, email=ol.smtp_of(entry))
    manager = None
    with contextlib.suppress(pywintypes.com_error):
        manager = user.GetExchangeUserManager()
    return Person(
        name=user.Name,
        email=get(user, "PrimarySmtpAddress", "") or "",
        title=get(user, "JobTitle", "") or "",
        department=get(user, "Department", "") or "",
        office=get(user, "OfficeLocation", "") or "",
        phone=get(user, "BusinessTelephoneNumber", "") or "",
        mobile=get(user, "MobileTelephoneNumber", "") or "",
        manager=(get(manager, "Name", "") or "") if manager is not None else "",
    )


class PeopleService:
    """People: resolve colleagues in the address book and report their availability / out-of-office."""

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook session of `ctx`."""
        self._session = ctx.outlook

    def lookup(self, query: str, *, availability_days: int = 0) -> tuple[Person, Availability | None]:
        """Contact card, plus availability when `availability_days` > 0."""

        def run(ol: Outlook) -> tuple[Person, Availability | None]:
            recipient = ol.recipient(query)
            person = _person(ol, recipient)
            if availability_days <= 0:
                return person, None
            at = now()
            return person, summarize(
                recipient_blocks(recipient, at, _FREEBUSY_MINUTES, availability_days), at
            )

        return self._session.run(run)
