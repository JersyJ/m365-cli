from __future__ import annotations

from m365_cli.core.errors import InvalidInputError
from m365_cli.features.mail.models import ComposeMode


def check_compose(mode: ComposeMode, *, ref: str | None, to: list[str] | None, subject: str | None) -> None:
    """Check that the compose arguments fit the mode (shared by all backends).

    Raises:
        InvalidInputError: An argument the mode needs is missing.
    """
    if mode == "new" and not (to and subject is not None):
        raise InvalidInputError("A new mail needs 'to' and 'subject'.")
    if mode != "new" and not ref:
        raise InvalidInputError(f"Mode '{mode}' needs mail_ref (the original mail).")
    if mode == "forward" and not to:
        raise InvalidInputError("Forward needs 'to'.")
