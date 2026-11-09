from __future__ import annotations

from m365_cli.core.timeutil import to_dasl_utc
from m365_cli.features.mail.models import MailQuery

_SUBJECT = '"urn:schemas:httpmail:subject"'
_FROM_NAME = '"urn:schemas:httpmail:fromname"'
_FROM_EMAIL = '"urn:schemas:httpmail:fromemail"'
_BODY = '"urn:schemas:httpmail:textdescription"'
_RECEIVED = '"urn:schemas:httpmail:datereceived"'
_READ = '"urn:schemas:httpmail:read"'
_HAS_ATTACHMENT = '"urn:schemas:httpmail:hasattachment"'
_FLAG_STATUS = '"http://schemas.microsoft.com/mapi/proptag/0x10900003"'


def _quote(value: str) -> str:
    return value.replace("'", "''")


def _like(field: str, value: str) -> str:
    return f"{field} LIKE '%{_quote(value)}%'"


def build_filter(q: MailQuery) -> str | None:
    """Return a DASL filter (without the '@SQL=' prefix), or None when the query has no conditions.

    DASL compares dates in UTC with ISO formatting, which avoids the locale-dependent parsing of
    Outlook's Jet syntax (e.g. on Czech Windows '2026-11-09' is read as 11 September).
    """
    parts: list[str] = []
    if q.since:
        parts.append(f"{_RECEIVED} >= '{to_dasl_utc(q.since)}'")
    if q.until:
        parts.append(f"{_RECEIVED} < '{to_dasl_utc(q.until)}'")
    if q.unread_only:
        parts.append(f"{_READ} = 0")
    if q.sender:
        parts.append(f"({_like(_FROM_NAME, q.sender)} OR {_like(_FROM_EMAIL, q.sender)})")
    if q.subject:
        parts.append(_like(_SUBJECT, q.subject))
    if q.text:
        parts.append(f"({_like(_SUBJECT, q.text)} OR {_like(_BODY, q.text)})")
    if q.has_attachments is not None:
        parts.append(f"{_HAS_ATTACHMENT} = {int(q.has_attachments)}")
    if q.flagged is not None:
        parts.append(
            f"{_FLAG_STATUS} = 2" if q.flagged else f"({_FLAG_STATUS} <> 2 OR {_FLAG_STATUS} IS NULL)"
        )
    return " AND ".join(parts) or None
