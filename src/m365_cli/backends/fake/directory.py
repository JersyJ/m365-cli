from __future__ import annotations

from m365_cli.backends.fake.state import FakePerson, FakeState
from m365_cli.core.errors import NotFoundError, UnresolvedRecipientsError

_FOLDER_ALIASES = {"trash": "deleted"}


def resolve_person(state: FakeState, query: str) -> FakePerson:
    """The one person whose name or email contains `query` (exact matches win).

    An unknown address containing "@" becomes an external person, as in Outlook.

    Raises:
        UnresolvedRecipientsError: No match, or more than one.
    """
    q = query.strip().lower()
    everyone = [state.me, *state.people]
    exact = [p for p in everyone if q in (p.name.lower(), p.email.lower())]
    matches = exact or [p for p in everyone if q in p.name.lower() or q in p.email.lower()]
    if len(matches) == 1:
        return matches[0]
    if not matches and "@" in q:
        return FakePerson(name=query.strip(), email=query.strip(), blocks=None)
    raise UnresolvedRecipientsError([query])


def label(person: FakePerson) -> str:
    """Return "Name <email>", or just the email when the name is the email."""
    return person.email if person.name == person.email else f"{person.name} <{person.email}>"


def resolve_mailbox(state: FakeState, mailbox: str | None) -> str:
    """Name of a shared mailbox in the state; '' for the user's own mailbox.

    Raises:
        NotFoundError: The mailbox is unknown.
    """
    if not mailbox or mailbox.strip().lower() in (state.me.name.lower(), state.me.email.lower()):
        return ""
    key = mailbox.strip().lower()
    matches = [m for m in state.mailboxes if m.lower() == key] or [
        m for m in state.mailboxes if key in m.lower()
    ]
    if len(matches) == 1:
        return matches[0]
    raise NotFoundError(
        f"No access to mailbox {mailbox!r}. Known: {', '.join([state.me.email, *state.mailboxes])}."
    )


def resolve_folder(state: FakeState, path: str) -> str:
    """Folder name in the state matching `path` (case-insensitive; 'trash' means 'deleted').

    Raises:
        NotFoundError: No such folder.
    """
    key = path.strip().strip("/").lower()
    key = _FOLDER_ALIASES.get(key, key)
    for folder in state.folders:
        if folder.lower() == key:
            return folder
    raise NotFoundError(f"Folder not found: {path!r}. Folders: {', '.join(state.folders)}.")
