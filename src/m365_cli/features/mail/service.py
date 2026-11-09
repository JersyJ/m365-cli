from __future__ import annotations

import re
from pathlib import Path

from m365_cli.backends.outlook import Com, Outlook, iterate
from m365_cli.backends.outlook.constants import OlFlagStatus, OlFolder, OlItemType, OlRecipientType
from m365_cli.backends.outlook.items import attachments_of, get, kind_of
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError
from m365_cli.core.paging import Page, take_page
from m365_cli.core.paths import to_windows_path
from m365_cli.core.text import clean_body, one_line, plain_to_html, shorten_body, with_signoff
from m365_cli.core.timeutil import from_com
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
from m365_cli.features.mail.query import build_filter
from m365_cli.features.mail.rules import check_compose

_BODY_START = re.compile(r"<body[^>]*>", re.I)
_PREVIEW_CHARS = 160
_FLAGS = {"set": OlFlagStatus.MARKED, "clear": OlFlagStatus.NONE, "complete": OlFlagStatus.COMPLETE}
_HTML_FORMAT = 2


class MailService:
    """Mail operations on classic Outlook. Every public method is safe to call from any thread.

    Sending is always two-step: `compose()` saves a draft (in Outlook's Drafts folder) and returns a
    preview; `send()` sends a draft by ref. Front-ends put the human confirmation between the two.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook session, refs, send log and settings of `ctx`."""
        self._session = ctx.outlook
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first."""

        def run(ol: Outlook) -> Page[MailSummary]:
            folder = ol.folder(query.folder)
            items = folder.Items
            if dasl := build_filter(query):
                items = items.Restrict("@SQL=" + dasl)
            is_sent = folder.EntryID == ol.default_folder(OlFolder.SENT).EntryID
            items.Sort("[SentOn]" if is_sent else "[ReceivedTime]", True)
            found, more = take_page(iterate(items), offset=offset, limit=query.limit)
            summaries = [
                self._summary(ol, it, folder.StoreID, is_sent=is_sent, preview=preview) for it in found
            ]
            return Page(items=summaries, offset=offset, has_more=more)

        return self._session.run(run)

    def _summary(
        self, ol: Outlook, item: Com, store_id: str, *, is_sent: bool = False, preview: bool = False
    ) -> MailSummary:
        who = str(get(item, "To", "") if is_sent else get(item, "SenderName", ""))
        return MailSummary(
            ref=self._refs.put(ol.backend_id(item, KIND, store_id)),
            received=from_com(get(item, "SentOn" if is_sent else "ReceivedTime")),
            sender=("to " + who) if is_sent else who,
            subject=str(get(item, "Subject", "") or "(no subject)"),
            unread=bool(get(item, "UnRead", False)),
            has_attachments=any(not a.inline for a in attachments_of(item)),
            flagged=get(item, "FlagStatus", 0) == OlFlagStatus.MARKED,
            kind=kind_of(item),
            preview=one_line(clean_body(str(get(item, "Body", "") or "")), _PREVIEW_CHARS) if preview else "",
        )

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        bid = self._refs.get(ref, kinds=KIND)

        def run(ol: Outlook) -> MailMessage:
            item = ol.item(bid)
            body, quoted_cut, truncated = shorten_body(
                clean_body(str(get(item, "Body", "") or "")), full=full, limit=self._settings.body_max_chars
            )
            return MailMessage(
                **self._summary(ol, item, bid.container_id).model_dump(),
                sender_email=ol.smtp_of(get(item, "Sender")),
                to=str(get(item, "To", "")),
                cc=str(get(item, "CC", "")),
                categories=str(get(item, "Categories", "") or ""),
                attachments=[
                    Attachment(index=a.index, name=a.name, size=a.size, inline=a.inline)
                    for a in attachments_of(item)
                ],
                body=body,
                quoted_cut=quoted_cut,
                truncated=truncated,
            )

        return self._session.run(run)

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
        """Save a draft (in Drafts) and return its preview. Never sends.

        Raises:
            InvalidInputError: Arguments don't fit the mode, or an attachment file is missing.
            UnresolvedRecipientsError: A recipient is unknown; nothing is kept.
        """
        check_compose(mode, ref=ref, to=to, subject=subject)
        files = [to_windows_path(p) for p in attachments or []]
        for f in files:
            if not Path(f).is_file():
                raise InvalidInputError(f"Attachment not found: {f}")
        text = with_signoff(body, self._settings.signoff)
        source = self._refs.get(ref, kinds=KIND) if ref else None

        def run(ol: Outlook) -> Draft:
            if source is None:
                item = ol.app.CreateItem(OlItemType.MAIL)
                if html:
                    item.HTMLBody = text
                else:
                    item.Body = text
            else:
                original = ol.item(source)
                actions = {
                    "reply": original.Reply,
                    "reply_all": original.ReplyAll,
                    "forward": original.Forward,
                }
                item = actions[mode]()
                _prepend_text(item, text, html=html)
            if subject is not None:
                item.Subject = subject
            groups = ((to, OlRecipientType.TO), (cc, OlRecipientType.CC), (bcc, OlRecipientType.BCC))
            for addresses, rtype in groups:
                for address in addresses or []:
                    item.Recipients.Add(address).Type = rtype
            for f in files:
                item.Attachments.Add(f)
            try:
                ol.resolve_all(item)
                if item.Recipients.Count == 0:
                    raise InvalidInputError("The draft has no recipients.")
            except Exception:
                item.Close(1)
                raise
            item.Save()
            return Draft(
                ref=self._refs.put(ol.backend_id(item, KIND, ol.default_store_id)),
                mode=mode,
                to=_labels(ol, item, OlRecipientType.TO),
                cc=_labels(ol, item, OlRecipientType.CC),
                bcc=_labels(ol, item, OlRecipientType.BCC),
                subject=str(item.Subject),
                text=text,
                includes_original=source is not None,
                attachments=[a.FileName for a in item.Attachments],
            )

        return self._session.run(run)

    def send(self, ref: str) -> str:
        """Send a draft (rate-limited and logged).

        Raises:
            InvalidInputError: The item is not a draft.
            RateLimitError: Too many sends in the last minute.
        """
        bid = self._refs.get(ref, kinds=KIND)
        self._audit.guard("mail_send")

        def run(ol: Outlook) -> tuple[str, str]:
            item = ol.item(bid)
            if get(item, "Sent", False):
                raise InvalidInputError(f"{ref} is not a draft (already sent/received).")
            ol.resolve_all(item)
            subject, recipients = str(item.Subject), "; ".join(r.Name for r in item.Recipients)
            item.Send()
            return subject, recipients

        subject, recipients = self._session.run(run)
        self._audit.record("mail_send", f"'{subject}' to {recipients}")
        return f"Sent '{subject}' to {recipients}."

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
        bid = self._refs.get(ref, kinds=KIND)

        def run(ol: Outlook) -> str:
            item = ol.item(bid)
            done: list[str] = []
            if read is not None:
                item.UnRead = not read
                done.append("marked read" if read else "marked unread")
            if flag is not None:
                item.FlagStatus = _FLAGS[flag]
                done.append(f"flag {flag}")
            if categories is not None:
                item.Categories = ", ".join(categories)
                done.append("categories: " + (", ".join(categories) or "none"))
            if done:
                item.Save()
            if delete:
                item.Delete()
                return "Deleted (moved to Deleted Items)."
            if move_to:
                target = ol.folder(move_to)
                moved = item.Move(target)
                new_ref = self._refs.put(ol.backend_id(moved, KIND, target.StoreID))
                done.append(f"moved to {target.Name} (new ref {new_ref})")
            if not done:
                raise InvalidInputError("Nothing to update.")
            return "; ".join(done).capitalize() + "."

        return self._session.run(run)

    def attachments(self, ref: str, *, save_to: str | None = None, include_inline: bool = False) -> list[str]:
        """List the attachments, or save them into directory `save_to`."""
        bid = self._refs.get(ref, kinds=KIND)
        target = Path(to_windows_path(save_to)) if save_to else None

        def run(ol: Outlook) -> list[str]:
            item = ol.item(bid)
            infos = [a for a in attachments_of(item) if include_inline or not a.inline]
            if target is None:
                return [f"{a.index}. {a.name} ({max(1, a.size // 1024)} KB)" for a in infos]
            target.mkdir(parents=True, exist_ok=True)
            saved = []
            for a in infos:
                path = target / a.name
                item.Attachments.Item(a.index).SaveAsFile(str(path))
                saved.append(str(path))
            return saved

        return self._session.run(run)


def _labels(ol: Outlook, item: Com, rtype: OlRecipientType) -> list[str]:
    labels = []
    for r in item.Recipients:
        if r.Type == rtype:
            smtp = ol.smtp_of(r.AddressEntry) if r.Resolved else ""
            labels.append(f"{r.Name} <{smtp}>" if smtp and smtp.lower() != r.Name.lower() else r.Name)
    return labels


def _prepend_text(item: Com, text: str, *, html: bool) -> None:
    """Insert our text above the quoted original, preserving the original's formatting."""
    if item.BodyFormat == _HTML_FORMAT:
        fragment = text if html else plain_to_html(text)
        body = item.HTMLBody
        m = _BODY_START.search(body)
        at = m.end() if m else 0
        item.HTMLBody = f"{body[:at]}<div>{fragment}</div><br>{body[at:]}"
    else:
        item.Body = f"{text}\n\n{item.Body}"
