from __future__ import annotations

import re
from collections import Counter, defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime

from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.features.mail.models import ComposeMode, DigestGroup, FlagAction


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


ALL_FOLDERS = "all"
"""Folder name that means: search every mail folder."""


def pick_attachment(options: list[tuple[int, str]], which: str) -> tuple[int, str]:
    """Choose an attachment by 1-based index, exact file name, or a unique part of the name.

    Raises:
        NotFoundError: No attachment, or more than one, matches.
    """
    key = which.strip()
    if key.isdigit():
        found = [o for o in options if o[0] == int(key)]
    else:
        found = [o for o in options if o[1].lower() == key.lower()] or [
            o for o in options if key.lower() in o[1].lower()
        ]
    if len(found) == 1:
        return found[0]
    names = ", ".join(f"{i}. {n}" for i, n in options) or "none"
    raise NotFoundError(f"No unique attachment {which!r}. Attachments: {names}.")


_SUBJECT_PREFIX = re.compile(r"^\s*(?:(?:re|fw|fwd|aw|wg|sv|vs|odp|tr|r)\s*(?:\[\d+\])?\s*:\s*)+", re.I)


def normalize_subject(subject: str) -> str:
    """Subject without reply/forward prefixes (RE:, FW:, AW:, SV:, …), for grouping and threads."""
    return _SUBJECT_PREFIX.sub("", subject or "").strip()


@dataclass(frozen=True, slots=True)
class DigestRow:
    """The few fields of one mail that the digest needs."""

    sender: str
    subject: str
    unread: bool
    received: datetime | None
    ref: Callable[[], str]


def group_digest(rows: Iterable[DigestRow], *, top: int = 3) -> list[DigestGroup]:
    """Group mails by sender, biggest groups first; newest ref and top subjects per group (pure)."""
    groups: dict[str, list[DigestRow]] = defaultdict(list)
    for row in rows:
        groups[row.sender or "(unknown sender)"].append(row)
    result = []
    for sender, items in groups.items():
        newest = max(items, key=lambda r: r.received or datetime.min)
        subjects = Counter(normalize_subject(r.subject) or "(no subject)" for r in items)
        result.append(
            DigestGroup(
                sender=sender,
                count=len(items),
                unread=sum(r.unread for r in items),
                newest=newest.received,
                newest_ref=newest.ref(),
                top_subjects=subjects.most_common(top),
            )
        )
    return sorted(result, key=lambda g: (-g.count, g.sender.lower()))


def is_answered(
    sent: datetime, subject: str, me: set[str], later: Iterable[tuple[str, str, datetime]]
) -> bool:
    """True when a later mail from someone else has the same normalized subject (pure).

    `later` holds (sender, subject, received) of candidate replies; `me` holds the user's names/addresses.
    """
    topic = normalize_subject(subject).lower()
    return any(
        received > sent and sender.lower() not in me and normalize_subject(subj).lower() == topic
        for sender, subj, received in later
    )


def update_summary(
    count: int,
    *,
    read: bool | None,
    flag: FlagAction | None,
    categories: list[str] | None,
    moved_to: str | None,
    moved_refs: list[str],
    delete: bool,
) -> str:
    """Short result text of a (bulk) update, shared by all backends."""
    done = []
    if read is not None:
        done.append("marked read" if read else "marked unread")
    if flag is not None:
        done.append(f"flag {flag}")
    if categories is not None:
        done.append("categories: " + (", ".join(categories) or "none"))
    if delete:
        done.append("moved to Deleted Items")
    elif moved_to:
        refs = f" (new ref {moved_refs[0]})" if len(moved_refs) == 1 else ""
        done.append(f"moved to {moved_to}{refs}")
    return f"{count} mail(s): " + "; ".join(done) + "."
