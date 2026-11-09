from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

from m365_cli.features.people.models import Availability, Person

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext


class PeopleApi(Protocol):
    """People operations the CLI and MCP tools use, independent of the backend."""

    def lookup(self, query: str, *, availability_days: int = 0) -> tuple[Person, Availability | None]:
        """Resolve a colleague by name, part of a name, or email.

        With `availability_days` > 0 also their availability for that many days.

        Raises:
            UnresolvedRecipientsError: No unique match.
        """
        ...


def people_api(ctx: AppContext) -> PeopleApi:
    """People operations of the active backend."""
    if ctx.is_fake:
        from m365_cli.features.people.fake import FakePeopleService

        return FakePeopleService(ctx)
    from m365_cli.features.people.service import PeopleService

    return PeopleService(ctx)
