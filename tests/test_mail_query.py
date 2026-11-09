from datetime import datetime

from m365_cli.features.mail.models import MailQuery
from m365_cli.features.mail.query import build_filter


def test_empty_query_has_no_filter() -> None:
    assert build_filter(MailQuery()) is None


def test_filter_combines_conditions_and_escapes_quotes() -> None:
    f = build_filter(MailQuery(unread_only=True, sender="O'Neil", subject="report", has_attachments=True))
    assert f is not None
    assert '"urn:schemas:httpmail:read" = 0' in f
    assert "LIKE '%O''Neil%'" in f
    assert "\"urn:schemas:httpmail:subject\" LIKE '%report%'" in f
    assert '"urn:schemas:httpmail:hasattachment" = 1' in f
    assert f.count(" AND ") == 3


def test_since_uses_iso_utc() -> None:
    f = build_filter(MailQuery(since=datetime(2026, 11, 9, 8, 0)))
    assert f is not None and f.startswith('"urn:schemas:httpmail:datereceived" >= \'2026-11-09 ')
