from datetime import datetime

import pytest

from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, UnresolvedRecipientsError
from m365_cli.core.timeutil import now
from m365_cli.features.calendar.api import calendar_api
from m365_cli.features.mail.api import mail_api
from m365_cli.features.mail.models import MailQuery
from m365_cli.features.people.api import people_api


def test_fake_backend_freezes_the_clock(fake_ctx: AppContext) -> None:
    assert now() == datetime(2026, 11, 9, 10, 0)


def test_mail_find_filters_pages_and_reads(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    page = mail.find(MailQuery(unread_only=True, limit=2))
    assert [m.subject for m in page.items] == ["Migration of runners - status", "Weekly GitLab Meeting"]
    assert page.has_more
    assert mail.find(MailQuery(sender="petr")).items[0].subject == "Weekly GitLab Meeting"
    message = mail.read(page.items[0].ref)
    assert message.body.endswith("Jana") and message.quoted_cut


def test_reply_creates_draft_and_send_records_outbox(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    source = mail.find(MailQuery(subject="runner capacity")).items[0]
    draft = mail.compose("reply", ref=source.ref, body="Ready on Friday.")
    assert draft.subject == "RE: Question: GitLab runner capacity"
    assert draft.to == ["Petr Svoboda <petr.svoboda@example.com>"]
    assert fake_ctx.fake.state.outbox == []
    assert "Sent" in mail.send(draft.ref)
    assert fake_ctx.fake.state.outbox[0].recipients == ["Petr Svoboda"]
    with pytest.raises(InvalidInputError):
        mail.send(draft.ref)


def test_unknown_recipient_keeps_nothing(fake_ctx: AppContext) -> None:
    before = len(fake_ctx.fake.state.mails)
    with pytest.raises(UnresolvedRecipientsError):
        mail_api(fake_ctx).compose("new", to=["Nobody Known"], subject="x", body="y")
    assert len(fake_ctx.fake.state.mails) == before


def test_update_flags_and_moves(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    ref = mail.find(MailQuery(subject="Cloud Weekly")).items[0].ref
    mail.update(ref, flag="set", move_to="Inbox/Newsletters")
    moved = mail.find(MailQuery(folder="inbox/newsletters")).items
    assert moved[0].flagged and moved[0].ref == ref


def test_calendar_free_time_draft_and_respond(fake_ctx: AppContext) -> None:
    cal = calendar_api(fake_ctx)
    free = cal.find_free_time(
        ["Jana Nováková", "Petr Svoboda", "partner@other.example"],
        day=datetime(2026, 11, 10),
        duration_minutes=60,
        limit=3,
    )
    assert [s.start.hour for s in free.slots] == [8, 12, 14]
    assert "partner@other.example" in free.notes[0]
    draft = cal.draft(subject="Sync", start=free.slots[0].start, end=free.slots[0].end, required=["Jana"])
    assert draft.is_meeting and draft.required == ["Jana Nováková"]
    gitlab = next(
        e
        for e in cal.find(datetime(2026, 11, 10), datetime(2026, 11, 11)).items
        if e.subject == "Weekly GitLab Meeting"
    )
    cal.respond(gitlab.ref, "accept")
    assert cal.read(gitlab.ref).my_response == "accepted"
    assert [o.action for o in fake_ctx.fake.state.outbox] == ["cal_respond"]


def test_canceled_events_are_hidden_by_default(fake_ctx: AppContext) -> None:
    cal = calendar_api(fake_ctx)
    day = (datetime(2026, 11, 10), datetime(2026, 11, 11))
    assert all(not e.canceled for e in cal.find(*day).items)
    assert any(e.canceled for e in cal.find(*day, include_canceled=True).items)


def test_people_lookup_with_ooo(fake_ctx: AppContext) -> None:
    person, availability = people_api(fake_ctx).lookup("eva", availability_days=5)
    assert person.title == "Team Manager"
    assert availability is not None and availability.out_of_office
