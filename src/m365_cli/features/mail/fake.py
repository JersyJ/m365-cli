from __future__ import annotations

from pathlib import PureWindowsPath

from m365_cli.backends.fake.directory import label, resolve_folder, resolve_person
from m365_cli.backends.fake.state import FakeMail, FakeState, Outgoing, SavedFile
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.paging import Page, take_page
from m365_cli.core.paths import to_windows_path
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, one_line, shorten_body, with_signoff
from m365_cli.core.timeutil import now
from m365_cli.features.mail.api import KIND
from m365_cli.features.mail.models import (
    Attachment,
    ComposeMode,
    Draft,
    FlagAction,
    MailMessage,
    MailQuery,
    MailSummary,
)
from m365_cli.features.mail.rules import check_compose

_PREVIEW_CHARS = 160
_PREFIX = {"reply": "RE: ", "reply_all": "RE: ", "forward": "FW: "}


def _contains(value: str, part: str | None) -> bool:
    return part is None or part.lower() in value.lower()


def _matches(mail: FakeMail, q: MailQuery, folder: str) -> bool:
    checks = [
        mail.folder == folder,
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

    def _summary(self, mail: FakeMail, *, preview: bool = False) -> MailSummary:
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
        )

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first."""
        state = self._store.state
        folder = resolve_folder(state, query.folder)
        found = sorted(
            (m for m in state.mails if _matches(m, query, folder)), key=lambda m: m.received, reverse=True
        )
        items, more = take_page(found, offset=offset, limit=query.limit)
        return Page(items=[self._summary(m, preview=preview) for m in items], offset=offset, has_more=more)

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        mail = self._mail(self._store.state, ref)
        body, quoted_cut, truncated = shorten_body(
            clean_body(mail.body), full=full, limit=self._settings.body_max_chars
        )
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
        html: bool = False,
        attachments: list[str] | None = None,
    ) -> Draft:
        """Save a draft (in Drafts) and return its preview. Never sends."""
        check_compose(mode, ref=ref, to=to, subject=subject)
        del html
        text = with_signoff(body, self._settings.signoff)
        with self._store.edit() as state:
            source = self._mail(state, ref) if ref else None
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
                body=text + quoted,
                is_draft=True,
                draft_mode=mode,
                source_id=source.id if source else None,
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
            includes_original=source is not None,
            attachments=files,
        )

    def send(self, ref: str) -> str:
        """Pretend to send a draft: move it to Sent and record it in the outbox (rate-limited, logged)."""
        self._audit.guard("mail_send")
        with self._store.edit() as state:
            mail = self._mail(state, ref)
            if not mail.is_draft:
                raise InvalidInputError(f"{ref} is not a draft (already sent/received).")
            mail.is_draft, mail.folder, mail.received = False, "sent", now()
            recipients = [*mail.to, *mail.cc, *mail.bcc]
            state.outbox.append(
                Outgoing(
                    at=now(), action="mail_send", item_id=mail.id, summary=mail.subject, recipients=recipients
                )
            )
        summary = f"'{mail.subject}' to {'; '.join(recipients)}"
        self._audit.record("mail_send", summary)
        return f"Sent {summary}."

    def update(
        self,
        ref: str,
        *,
        read: bool | None = None,
        flag: FlagAction | None = None,
        categories: list[str] | None = None,
        move_to: str | None = None,
        delete: bool = False,
    ) -> str:
        """Mark read/unread, flag, set categories, move or delete; returns a short result."""
        with self._store.edit() as state:
            mail = self._mail(state, ref)
            done: list[str] = []
            if read is not None:
                mail.unread = not read
                done.append("marked read" if read else "marked unread")
            if flag is not None:
                mail.flag = "none" if flag == "clear" else flag
                done.append(f"flag {flag}")
            if categories is not None:
                mail.categories = categories
                done.append("categories: " + (", ".join(categories) or "none"))
            if delete:
                mail.folder = "deleted"
                return "Deleted (moved to Deleted Items)."
            if move_to:
                mail.folder = resolve_folder(state, move_to)
                done.append(f"moved to {mail.folder} (ref {ref} is still valid)")
            if not done:
                raise InvalidInputError("Nothing to update.")
        return "; ".join(done).capitalize() + "."

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
