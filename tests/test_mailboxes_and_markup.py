from datetime import datetime

import pytest

from m365_cli.backends.fake.fixtures import SHARED_MAILBOX
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.markup import (
    MailStyle,
    body_fragment,
    decode_html_file,
    parse_style,
    styled,
    to_html,
    to_text,
)
from m365_cli.features.calendar.api import calendar_api
from m365_cli.features.mail import render
from m365_cli.features.mail.api import mail_api
from m365_cli.features.mail.models import MailMessage, MailQuery
from tests.test_mcp_server import call

TOMORROW = (datetime(2026, 11, 10), datetime(2026, 11, 11))


def test_parse_outlook_compose_style() -> None:
    css = 'span.Compose {mso-ansi-font-size:12.0pt; font-family:"Aptos",sans-serif; color:#1F3864;}'
    style = parse_style(css)
    assert style == MailStyle(font_family='"Aptos",sans-serif', font_size_pt=12.0, color="#1F3864")
    wrapped = styled("<p>x</p>", style)
    assert wrapped.startswith('<div style="font-family:&quot;Aptos&quot;,sans-serif;font-size:12pt')
    assert parse_style("color:windowtext").color == ""


def test_markdown_html_and_text_bodies() -> None:
    html = to_html("Hi,\n**bold** and [link](https://example.com)\n\n- a\n- b", "markdown")
    assert (
        "<strong>bold</strong>" in html and 'href="https://example.com"' in html and html.count("<li>") == 2
    )
    assert "<br" in html
    assert to_html("<b>x</b>", "markdown") == "<p>&lt;b&gt;x&lt;/b&gt;</p>"
    assert to_html("a < b", "text") == "a &lt; b"
    assert to_text("<p>One</p><p>Two &amp; three</p>", "html") == "One\nTwo & three"


def test_signature_file_decoding() -> None:
    raw = '<html><head><meta charset="windows-1252"></head><body><p>Alex Testovsk\xfd</p></body></html>'
    assert body_fragment(decode_html_file(raw.encode("windows-1252"))) == "<p>Alex Testovský</p>"


def test_draft_gets_signature_and_format(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    new = mail.compose("new", to=["Jana"], subject="Plan", body="**Go**")
    assert new.signature.startswith("new") and new.body_format == "markdown"
    stored = next(m for m in fake_ctx.fake.state.mails if m.id == fake_ctx.refs.get(new.ref).item_id)
    assert stored.body.endswith("Example Corp")
    reply = mail.compose("reply", ref=mail.find(MailQuery(sender="petr")).items[0].ref, body="Ok")
    assert reply.signature.startswith("reply")


def test_shared_mailbox_find_and_reply_on_behalf(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    assert (SHARED_MAILBOX, "shared") in mail.mailboxes()
    own = mail.find(MailQuery(subject="prod subscription")).items
    shared = mail.find(MailQuery(mailbox="team", subject="prod subscription")).items
    assert own == [] and len(shared) == 1
    draft = mail.compose("reply", ref=shared[0].ref, body="Granted.")
    assert draft.from_mailbox == f"{SHARED_MAILBOX} (on behalf)"
    assert "From: Team Shared (on behalf)" in render.draft(draft)
    with pytest.raises(NotFoundError, match="Known:"):
        mail.find(MailQuery(mailbox="nonexistent"))
    assert "Team Shared (shared)" in call(fake_ctx, "mail_mailboxes")


def test_shared_calendar_and_no_access(fake_ctx: AppContext) -> None:
    cal = calendar_api(fake_ctx)
    events = cal.find(*TOMORROW, owner="Jana").items
    assert [e.subject for e in events] == ["Workshop: landing zones"]
    assert cal.read(events[0].ref).organizer == "Jana Nováková"
    with pytest.raises(InvalidInputError, match="not shared with you"):
        cal.find(*TOMORROW, owner="Petr")


def test_protected_mail_shows_note_instead_of_wrapper() -> None:
    message = MailMessage(
        ref="m:X",
        received=None,
        sender="A",
        subject="Secret",
        protected=True,
        body="This message is protected with Microsoft Information Protection.",
    )
    text = render.message(message, mark_untrusted=True)
    assert "Protected mail" in text and "This message is protected" not in text
    assert "[protected]" in render.summary_line(message)


def test_repeat_rules_and_recurring_draft(fake_ctx: AppContext) -> None:
    from m365_cli.features.calendar.models import Repeat
    from m365_cli.features.calendar.rules import check_repeat, describe_repeat

    monday = datetime(2026, 11, 16, 9)
    weekly = Repeat(every="week", days=["mon", "wed"], until=datetime(2026, 12, 31).date())
    assert describe_repeat(weekly, monday) == "every week on Mon, Wed until 2026-12-31"
    assert check_repeat(Repeat(every="week"), monday) == ["mon"]
    assert describe_repeat(Repeat(every="month", interval=2, count=6), monday) == "every 2 months, 6 times"
    with pytest.raises(InvalidInputError, match="either until or count"):
        check_repeat(Repeat(every="day", until=monday.date(), count=3), monday)
    with pytest.raises(InvalidInputError, match="only works with every='week'"):
        check_repeat(Repeat(every="day", days=["mon"]), monday)
    draft = calendar_api(fake_ctx).draft(
        subject="Weekly sync", start=monday, end=monday.replace(hour=10), required=["Jana"], repeat=weekly
    )
    assert draft.repeat == weekly
    assert "Repeats: every day, 5 times" in call(
        fake_ctx,
        "cal_draft",
        subject="Standup",
        start="2026-11-16 09:00",
        repeat={"every": "day", "count": 5},
    )
