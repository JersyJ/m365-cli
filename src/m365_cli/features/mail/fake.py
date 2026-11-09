from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import PureWindowsPath

from m365_cli.backends.fake.directory import label, resolve_folder, resolve_mailbox, resolve_person
from m365_cli.backends.fake.state import FakeMail, FakeState, Outgoing, SavedFile
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.markup import BodyFormat, to_text
from m365_cli.core.paging import Page, take_page
from m365_cli.core.paths import to_windows_path
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, one_line, shorten_body, truncate, with_signoff
from m365_cli.core.timeutil import fmt, now
from m365_cli.features.mail.api import KIND
from m365_cli.features.mail.models import (
    Attachment,
    AttachmentText,
    ComposeMode,
    DigestGroup,
    Draft,
    FlagAction,
    MailMessage,
    MailQuery,
    MailSummary,
    WaitingItem,
)
from m365_cli.features.mail.rules import (
    ALL_FOLDERS,
    DigestRow,
    check_compose,
    group_digest,
    is_answered,
    normalize_subject,
    pick_attachment,
    update_summary,
)

_PREVIEW_CHARS = 160
_THREAD_BODY_CHARS = 2000
_PREFIX = {"reply": "RE: ", "reply_all": "RE: ", "forward": "FW: "}
_NOT_SEARCHED = frozenset({"deleted", "junk", "drafts", "outbox"})


def _contains(value: str, part: str | None) -> bool:
    return part is None or part.lower() in value.lower()


def _matches(mail: FakeMail, q: MailQuery, folders: set[str]) -> bool:
    checks = [
        mail.folder in folders,
        not q.unread_only or mail.unread,
        q.since is None or mail.received >= q.since,
        q.until is None or mail.received < q.until,
        q.sender is None or _contains(mail.sender_name, q.sender) or _contains(mail.sender_email, q.sender),
        _contains(mail.subject, q.subject),
        q.text is None or _contains(mail.subject, q.text) or _contains(mail.body, q.text),
        q.has_attachments is None or q.has_attachments == any(not a.inline for a in mail.attachments),
        q.flagged is None or q.flagged == (mail.flag == "set"),
    ]
    return all(checks)


class FakeMailService:
    """Mail operations on the fake backend.

    Behaves like `MailService` for everything the tests and evaluations use, but never sends:
    `send()` records the mail in `FakeState.outbox`.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the fake store, refs, send log and settings of `ctx`."""
        self._store = ctx.fake
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def _ref(self, mail: FakeMail) -> str:
        return self._refs.put(BackendId(KIND, mail.id))

    def _mail(self, state: FakeState, ref: str) -> FakeMail:
        mail_id = self._refs.get(ref, kinds=KIND).item_id
        for mail in state.mails:
            if mail.id == mail_id:
                return mail
        raise NotFoundError("Mail no longer exists; list again.")

    def _summary(self, mail: FakeMail, *, preview: bool = False, show_folder: bool = False) -> MailSummary:
        is_sent = mail.folder == "sent"
        return MailSummary(
            ref=self._ref(mail),
            received=mail.received,
            sender=("to " + "; ".join(mail.to)) if is_sent else mail.sender_name,
            subject=mail.subject or "(no subject)",
            unread=mail.unread,
            has_attachments=any(not a.inline for a in mail.attachments),
            flagged=mail.flag == "set",
            kind=mail.kind,
            preview=one_line(clean_body(mail.body), _PREVIEW_CHARS) if preview else "",
            folder=mail.folder if show_folder else "",
            protected=mail.protected,
        )

    def _matching(self, state: FakeState, query: MailQuery) -> list[FakeMail]:
        mailbox = resolve_mailbox(state, query.mailbox)
        if query.folder.strip().lower() == ALL_FOLDERS:
            folders = {f for f in state.folders if f not in _NOT_SEARCHED}
        else:
            folders = {resolve_folder(state, query.folder)}
        found = (m for m in state.mails if m.mailbox == mailbox and _matches(m, query, folders))
        return sorted(found, key=lambda m: m.received, reverse=True)

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first. Folder "all" searches every mail folder."""
        everywhere = query.folder.strip().lower() == ALL_FOLDERS
        items, more = take_page(self._matching(self._store.state, query), offset=offset, limit=query.limit)
        summaries = [self._summary(m, preview=preview, show_folder=everywhere) for m in items]
        return Page(items=summaries, offset=offset, has_more=more)

    def select(self, query: MailQuery, *, max_items: int) -> list[str]:
        """Refs of up to `max_items` mails matching the query (cheap; for bulk updates)."""
        return [self._ref(m) for m in self._matching(self._store.state, query)[:max_items]]

    def digest(self, query: MailQuery, *, max_items: int = 300) -> list[DigestGroup]:
        """Mails matching the query, grouped by sender (at most `max_items` mails are counted)."""
        rows = [
            DigestRow(
                sender=m.sender_name,
                subject=m.subject,
                unread=m.unread,
                received=m.received,
                ref=lambda mail=m: self._ref(mail),
            )
            for m in self._matching(self._store.state, query)[:max_items]
        ]
        return group_digest(rows)

    def _message(self, mail: FakeMail, *, full: bool, limit: int) -> MailMessage:
        body, quoted_cut, truncated = shorten_body(clean_body(mail.body), full=full, limit=limit)
        return MailMessage(
            **self._summary(mail).model_dump(),
            sender_email=mail.sender_email,
            to="; ".join(mail.to),
            cc="; ".join(mail.cc),
            categories=", ".join(mail.categories),
            attachments=[
                Attachment(index=i, name=a.name, size=a.size, inline=a.inline)
                for i, a in enumerate(mail.attachments, start=1)
            ],
            body=body,
            quoted_cut=quoted_cut,
            truncated=truncated,
        )

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        return self._message(
            self._mail(self._store.state, ref), full=full, limit=self._settings.body_max_chars
        )

    def thread(self, ref: str, *, full: bool = False) -> list[MailMessage]:
        """The whole conversation (same normalized subject, any folder but drafts), oldest first."""
        state = self._store.state
        topic = normalize_subject(self._mail(state, ref).subject).lower()
        members = [m for m in state.mails if not m.is_draft and normalize_subject(m.subject).lower() == topic]
        return [
            self._message(m, full=full, limit=_THREAD_BODY_CHARS)
            for m in sorted(members, key=lambda m: m.received)
        ]

    def mailboxes(self) -> list[tuple[str, str]]:
        """The user's mailbox plus the shared ones of the fake state."""
        state = self._store.state
        return [(state.me.email, "primary"), *((m, "shared") for m in state.mailboxes)]

    def waiting(
        self, *, days: int = 3, within_days: int = 14, limit: int = 20, mailbox: str | None = None
    ) -> list[WaitingItem]:
        """Sent mails from the last `within_days` days, older than `days`, with no reply yet."""
        if days < 0 or within_days < days:
            raise InvalidInputError("Need 0 <= days <= within_days.")
        state = self._store.state
        me = {state.me.name.lower(), state.me.email.lower()}
        newest_allowed, oldest_allowed = now() - timedelta(days=days), now() - timedelta(days=within_days)
        later = [(m.sender_name, m.subject, m.received) for m in state.mails if m.folder != "sent"]
        box = resolve_mailbox(state, mailbox)
        sent = sorted(
            (m for m in state.mails if m.folder == "sent" and m.mailbox == box),
            key=lambda m: m.received,
            reverse=True,
        )
        result = []
        for mail in sent:
            if not oldest_allowed <= mail.received <= newest_allowed or mail.kind == "meeting":
                continue
            if {n.lower() for n in mail.to} <= me or is_answered(mail.received, mail.subject, me, later):
                continue
            result.append(
                WaitingItem(
                    ref=self._ref(mail),
                    sent=mail.received,
                    to="; ".join(mail.to),
                    subject=mail.subject,
                    days=(now() - mail.received).days,
                )
            )
        return result[:limit]

    def compose(
        self,
        mode: ComposeMode,
        *,
        body: str,
        ref: str | None = None,
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        subject: str | None = None,
        body_format: BodyFormat = "markdown",
        attachments: list[str] | None = None,
        from_mailbox: str | None = None,
    ) -> Draft:
        """Save a draft (in Drafts) and return its preview. Never sends.

        The signature comes from `FakeState.signatures` ("new" / "reply"), like Outlook's defaults.
        """
        check_compose(mode, ref=ref, to=to, subject=subject)
        text = with_signoff(body, self._settings.signoff)
        kind = "new" if ref is None else "reply"
        setting = getattr(self._settings, f"signature_{kind}")
        with self._store.edit() as state:
            source = self._mail(state, ref) if ref else None
            signature = "" if setting.lower() == "none" else state.signatures.get(kind, "")
            sender = (
                resolve_mailbox(state, from_mailbox) if from_mailbox else (source.mailbox if source else "")
            )
            to_names, cc_names = list(to or []), list(cc or [])
            files: list[str] = [PureWindowsPath(to_windows_path(p)).name for p in attachments or []]
            if source is not None:
                if mode in ("reply", "reply_all"):
                    to_names = [source.sender_email or source.sender_name, *to_names]
                if mode == "reply_all":
                    others = [n for n in [*source.to, *source.cc] if n not in (state.me.name, state.me.email)]
                    cc_names = [*others, *cc_names]
                if mode == "forward":
                    files = [a.name for a in source.attachments if not a.inline] + files
            to_people = [resolve_person(state, n) for n in to_names]
            cc_people = [resolve_person(state, n) for n in cc_names]
            bcc_people = [resolve_person(state, n) for n in bcc or []]
            if not to_people:
                raise InvalidInputError("The draft has no recipients.")
            title = subject if subject is not None else _PREFIX[mode] + (source.subject if source else "")
            quoted = (
                f"\n\nFrom: {source.sender_name}\nSubject: {source.subject}\n\n{source.body}"
                if source
                else ""
            )
            draft = FakeMail(
                id=state.new_id("d"),
                folder="drafts",
                received=now(),
                sender_name=state.me.name,
                sender_email=state.me.email,
                to=[p.name for p in to_people],
                cc=[p.name for p in cc_people],
                bcc=[p.name for p in bcc_people],
                subject=title,
                body=to_text(text, body_format) + (f"\n\n{signature}" if signature else "") + quoted,
                is_draft=True,
                draft_mode=mode,
                source_id=source.id if source else None,
                from_mailbox=sender,
                attachments=[a for a in (source.attachments if source else []) if mode == "forward"],
            )
            state.mails.append(draft)
        return Draft(
            ref=self._ref(draft),
            mode=mode,
            to=[label(p) for p in to_people],
            cc=[label(p) for p in cc_people],
            bcc=[label(p) for p in bcc_people],
            subject=title,
            text=text,
            body_format=body_format,
            signature=f"{kind} ({len(signature.splitlines())} lines)" if signature else "",
            from_mailbox=f"{draft.from_mailbox} (on behalf)" if draft.from_mailbox else "",
            includes_original=source is not None,
            attachments=files,
        )

    def send(self, ref: str, *, send_at: datetime | None = None) -> str:
        """Pretend to send (or schedule) a draft: move it to Sent and record it in the outbox."""
        if send_at is not None and send_at <= now():
            raise InvalidInputError(f"send_at {fmt(send_at)} is in the past.")
        self._audit.guard("mail_send")
        with self._store.edit() as state:
            mail = self._mail(state, ref)
            if not mail.is_draft:
                raise InvalidInputError(f"{ref} is not a draft (already sent/received).")
            mail.is_draft, mail.folder, mail.received = False, "sent", send_at or now()
            recipients = [*mail.to, *mail.cc, *mail.bcc]
            state.outbox.append(
                Outgoing(
                    at=now(),
                    action="mail_send",
                    item_id=mail.id,
                    summary=mail.subject,
                    recipients=recipients,
                    send_at=send_at,
                )
            )
        when = f" (scheduled for {fmt(send_at)})" if send_at else ""
        summary = f"'{mail.subject}' to {'; '.join(recipients)}{when}"
        self._audit.record("mail_send", summary)
        return f"{'Scheduled' if send_at else 'Sent'} {summary}."

    def update(
        self,
        refs: list[str],
        *,
        read: bool | None = None,
        flag: FlagAction | None = None,
        categories: list[str] | None = None,
        move_to: str | None = None,
        delete: bool = False,
    ) -> str:
        """Mark read/unread, flag, set categories, move or delete one or more mails; returns a summary."""
        if not refs:
            raise InvalidInputError("No mails to update.")
        if read is None and flag is None and categories is None and not move_to and not delete:
            raise InvalidInputError("Nothing to update.")
        with self._store.edit() as state:
            target = resolve_folder(state, move_to) if move_to else None
            for ref in refs:
                mail = self._mail(state, ref)
                if read is not None:
                    mail.unread = not read
                if flag is not None:
                    mail.flag = "none" if flag == "clear" else flag
                if categories is not None:
                    mail.categories = categories
                if delete:
                    mail.folder = "deleted"
                elif target is not None:
                    mail.folder = target
        return update_summary(
            len(refs),
            read=read,
            flag=flag,
            categories=categories,
            moved_to=target,
            moved_refs=[],
            delete=delete,
        )

    def attachments(self, ref: str, *, save_to: str | None = None, include_inline: bool = False) -> list[str]:
        """List the attachments, or record them as saved into directory `save_to`."""
        state = self._store.state
        mail = self._mail(state, ref)
        infos = [(i, a) for i, a in enumerate(mail.attachments, start=1) if include_inline or not a.inline]
        if save_to is None:
            return [f"{i}. {a.name} ({max(1, a.size // 1024)} KB)" for i, a in infos]
        target = PureWindowsPath(to_windows_path(save_to))
        with self._store.edit() as editable:
            for _, a in infos:
                editable.saved_files.append(
                    SavedFile(mail_id=mail.id, name=a.name, path=str(target / a.name))
                )
        return [str(target / a.name) for _, a in infos]

    def attachment_text(self, ref: str, which: str) -> AttachmentText:
        """Text stored with the fake attachment, chosen by 1-based index or by file name."""
        mail = self._mail(self._store.state, ref)
        index, name = pick_attachment([(i, a.name) for i, a in enumerate(mail.attachments, start=1)], which)
        text, cut = truncate(mail.attachments[index - 1].text or "(no text)", self._settings.body_max_chars)
        return AttachmentText(name=name, text=text, truncated=cut)
