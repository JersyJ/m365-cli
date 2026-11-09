class M365Error(Exception):
    """Base class for expected, user-facing errors (bad input, missing item, backend unavailable).

    The CLI and the MCP server print the message of these errors; any other exception is a bug.
    """


class NotFoundError(M365Error):
    """A referenced item, folder or person does not exist (any more)."""


class InvalidInputError(M365Error):
    """Arguments are malformed or contradictory."""


class BackendUnavailableError(M365Error):
    """The backend (e.g. classic Outlook) cannot be reached or is not signed in."""


class RateLimitError(M365Error):
    """Too many sends in the last minute; nothing was sent."""


class UnresolvedRecipientsError(M365Error):
    """One or more recipients could not be resolved to an address; nothing was sent or saved."""

    def __init__(self, names: list[str]) -> None:
        """Store the names that could not be resolved."""
        super().__init__(
            "Could not resolve recipient(s): " + ", ".join(names) + ". Use full names or emails."
        )
        self.names = names
