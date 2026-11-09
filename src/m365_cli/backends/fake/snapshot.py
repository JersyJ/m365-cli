from __future__ import annotations

import itertools
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta

from m365_cli.backends.fake.fixtures import INJECTION_MAIL
from m365_cli.backends.fake.state import (
    FakeAttachment,
    FakeBlock,
    FakeEvent,
    FakeMail,
    FakePerson,
    FakeState,
    FlagState,
)
from m365_cli.backends.outlook import Com, Outlook, OutlookSession, iterate
from m365_cli.backends.outlook.constants import OlFlagStatus, OlFolder, OlMeetingStatus, OlResponseStatus
from m365_cli.backends.outlook.freebusy import recipient_blocks
from m365_cli.backends.outlook.items import attachments_of, get, kind_of, split_names
from m365_cli.core.text import clean_body, truncate
from m365_cli.core.timeutil import from_com, now, to_dasl_utc

_BODY_CHARS = 4000
_FREEBUSY_MINUTES = 30
_RESPONSES = {
    OlResponseStatus.ORGANIZED: "organizer",
    OlResponseStatus.TENTATIVE: "tentative",
    OlResponseStatus.ACCEPTED: "accepted",
    OlResponseStatus.DECLINED: "declined",
    OlResponseStatus.NOT_RESPONDED: "not responded",
}


@dataclass(frozen=True, slots=True)
class SnapshotOptions:
    """How much real data a snapshot copies."""

    mail_days: int = 14
    max_inbox: int = 100
    max_sent: int = 30
    max_per_subfolder: int = 10
    event_days_back: int = 7
    event_days_ahead: int = 14
    max_people: int = 40
    add_injection_mail: bool = False


def take_snapshot(session: OutlookSession, options: SnapshotOptions | None = None) -> FakeState:
    """Copy real data from classic Outlook into a fake state, so evaluations run on your real mail.

    Copies the mails of the last `mail_days` days (Inbox, its subfolders, Sent Items), the calendar
    from `event_days_back` days ago to `event_days_ahead` days ahead, and the colleagues that occur
    in them with their free/busy. Each step is a separate Outlook call, so none hits the time-out.

    The result contains private data: keep the file local and out of version control.
    """
    opts = options or SnapshotOptions()
    clock = now()
    me = session.run(_me)
    folders = session.run(_folders)
    events, event_keys = session.run(lambda ol: _events(ol, clock, opts))
    mails = session.run(lambda ol: _mails(ol, clock, opts, event_keys, events))
    people = _people(session, clock, opts, mails, events, me)
    if opts.add_injection_mail:
        mails.append(
            INJECTION_MAIL.model_copy(update={"received": clock - timedelta(hours=26), "to": [me.name]})
        )
    return FakeState(
        source="snapshot", clock=clock, me=me, folders=folders, mails=mails, events=events, people=people
    )


def _me(ol: Outlook) -> FakePerson:
    user = ol.ns.CurrentUser.AddressEntry.GetExchangeUser()
    return FakePerson(
        name=user.Name,
        email=user.PrimarySmtpAddress,
        title=user.JobTitle or "",
        department=user.Department or "",
    )


def _folders(ol: Outlook) -> list[str]:
    inbox = ol.default_folder(OlFolder.INBOX)
    subfolders = [f"Inbox/{f.Name}" for f in inbox.Folders]
    return ["inbox", "sent", "drafts", "deleted", "junk", *subfolders]


def _events(
    ol: Outlook, clock: datetime, opts: SnapshotOptions
) -> tuple[list[FakeEvent], dict[tuple[str, str], str]]:
    start = clock.replace(hour=0, minute=0, second=0) - timedelta(days=opts.event_days_back)
    end = start + timedelta(days=opts.event_days_back + opts.event_days_ahead + 1)
    events: list[FakeEvent] = []
    keys: dict[tuple[str, str], str] = {}
    for n, item in enumerate(iterate(ol.calendar_between(start, end)), start=1):
        event = _event(item, f"e{n}")
        events.append(event)
        keys[(item.EntryID, event.start.isoformat())] = event.id
    return events, keys


def _event(item: Com, event_id: str) -> FakeEvent:
    status = get(item, "MeetingStatus", 0)
    body = clean_body(str(get(item, "Body", "") or ""))
    start, end = from_com(item.Start) or now(), from_com(item.End) or now()
    join = next((w for w in body.split() if "teams.microsoft.com/l/meetup-join" in w), "")
    return FakeEvent(
        id=event_id,
        start=start,
        end=end,
        all_day=bool(get(item, "AllDayEvent", False)),
        subject=str(get(item, "Subject", "") or "(no subject)"),
        location=str(get(item, "Location", "") or ""),
        organizer=str(get(item, "Organizer", "") or ""),
        organizer_is_me=get(item, "ResponseStatus", 0) == OlResponseStatus.ORGANIZED,
        required=split_names(get(item, "RequiredAttendees", "")),
        optional=split_names(get(item, "OptionalAttendees", "")),
        body=truncate(body, _BODY_CHARS)[0],
        show_as=int(get(item, "BusyStatus", 2)),
        my_response=_RESPONSES.get(get(item, "ResponseStatus", 0), "")
        if status != OlMeetingStatus.NON_MEETING
        else "",
        recurring=bool(get(item, "IsRecurring", False)),
        canceled=status in (OlMeetingStatus.CANCELED, OlMeetingStatus.RECEIVED_AND_CANCELED),
        is_meeting=status != OlMeetingStatus.NON_MEETING,
        join_url=join.strip("<>"),
    )


def _mails(
    ol: Outlook,
    clock: datetime,
    opts: SnapshotOptions,
    event_keys: dict[tuple[str, str], str],
    events: list[FakeEvent],
) -> list[FakeMail]:
    since = to_dasl_utc(clock - timedelta(days=opts.mail_days))
    inbox = ol.default_folder(OlFolder.INBOX)
    sources: list[tuple[str, Com, int]] = [
        ("inbox", inbox, opts.max_inbox),
        ("sent", ol.default_folder(OlFolder.SENT), opts.max_sent),
    ]
    sources += [(f"Inbox/{f.Name}", f, opts.max_per_subfolder) for f in inbox.Folders]
    mails: list[FakeMail] = []
    counter = itertools.count(1)
    for folder_name, folder, cap in sources:
        items = folder.Items.Restrict(f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{since}'")
        items.Sort("[ReceivedTime]", True)
        for item in itertools.islice(iterate(items), cap):
            mails.append(_mail(item, f"m{next(counter)}", folder_name, event_keys, events))
    return mails


def _mail(
    item: Com, mail_id: str, folder: str, event_keys: dict[tuple[str, str], str], events: list[FakeEvent]
) -> FakeMail:
    kind = kind_of(item)
    event_id = None
    if kind == "meeting":
        appointment = get(item, "GetAssociatedAppointment", None)
        appt = appointment(False) if appointment else None
        if appt is not None:
            key = (appt.EntryID, (from_com(appt.Start) or now()).isoformat())
            event_id = event_keys.get(key)
            if event_id is None:
                event_id = f"e{len(events) + 1}"
                events.append(_event(appt, event_id))
                event_keys[key] = event_id
    flags: dict[int, FlagState] = {OlFlagStatus.MARKED: "set", OlFlagStatus.COMPLETE: "complete"}
    flag = flags.get(get(item, "FlagStatus", 0), "none")
    sender_email = ""
    sender = get(item, "Sender")
    if sender is not None:
        user = get(sender, "GetExchangeUser")
        exchange_user = user() if user and get(sender, "AddressEntryUserType", -1) in (0, 5) else None
        sender_email = exchange_user.PrimarySmtpAddress if exchange_user else str(get(sender, "Address", ""))
    return FakeMail(
        id=mail_id,
        folder=folder,
        received=from_com(get(item, "ReceivedTime")) or now(),
        sender_name=str(get(item, "SenderName", "") or ""),
        sender_email=sender_email,
        to=split_names(get(item, "To", "")),
        cc=split_names(get(item, "CC", "")),
        subject=str(get(item, "Subject", "") or "(no subject)"),
        body=truncate(clean_body(str(get(item, "Body", "") or "")), _BODY_CHARS)[0],
        unread=bool(get(item, "UnRead", False)),
        flag=flag,
        categories=[c.strip() for c in str(get(item, "Categories", "") or "").split(",") if c.strip()],
        attachments=[FakeAttachment(name=a.name, size=a.size, inline=a.inline) for a in attachments_of(item)],
        kind=kind,
        event_id=event_id,
    )


def _people(
    session: OutlookSession,
    clock: datetime,
    opts: SnapshotOptions,
    mails: list[FakeMail],
    events: list[FakeEvent],
    me: FakePerson,
) -> list[FakePerson]:
    counts: Counter[str] = Counter()
    for m in mails:
        counts.update([m.sender_email or m.sender_name, *m.to, *m.cc])
    for e in events:
        counts.update([e.organizer, *e.required, *e.optional])
    candidates = [n for n, _ in counts.most_common() if n and n not in (me.name, me.email)]
    people: list[FakePerson] = []
    seen: set[str] = set()
    for name in candidates:
        if len(people) >= opts.max_people:
            break
        person = session.run(lambda ol, n=name: _person(ol, n, clock))
        if person is not None and person.email.lower() not in seen:
            seen.add(person.email.lower())
            people.append(person)
    return people


def _person(ol: Outlook, query: str, clock: datetime) -> FakePerson | None:
    recipient = ol.ns.CreateRecipient(query)
    if not recipient.Resolve():
        return None
    user = get(recipient.AddressEntry, "GetExchangeUser")
    exchange_user = user() if user else None
    if exchange_user is None:
        return None
    manager = get(exchange_user, "GetExchangeUserManager")
    manager_user = manager() if manager else None
    blocks = recipient_blocks(recipient, clock, _FREEBUSY_MINUTES, 14)
    return FakePerson(
        name=exchange_user.Name,
        email=exchange_user.PrimarySmtpAddress or "",
        title=exchange_user.JobTitle or "",
        department=exchange_user.Department or "",
        office=exchange_user.OfficeLocation or "",
        phone=exchange_user.BusinessTelephoneNumber or "",
        mobile=exchange_user.MobileTelephoneNumber or "",
        manager=manager_user.Name if manager_user else "",
        blocks=None
        if blocks is None
        else [FakeBlock(start=b.start, end=b.end, status=int(b.status)) for b in blocks],
    )
