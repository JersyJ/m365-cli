from __future__ import annotations

import base64
import heapq
import itertools
import re
import tempfile
from collections.abc import Iterator
from datetime import datetime, timedelta
from pathlib import Path

import pywintypes

from m365_cli.backends.outlook import Com, Outlook, iterate
from m365_cli.backends.outlook.constants import (
    OL_DISCARD,
    OlFlagStatus,
    OlFolder,
    OlItemType,
    OlRecipientType,
)
from m365_cli.backends.outlook.items import attachments_of, get, is_protected, kind_of
from m365_cli.backends.outlook.profile import SignatureKind, mail_style, signature_for
from m365_cli.core.context import AppContext
from m365_cli.core.doctext import extract_text
from m365_cli.core.errors import InvalidInputError
from m365_cli.core.markup import BodyFormat, Signature, html_document, styled, to_html, to_text
from m365_cli.core.paging import Page, take_page
from m365_cli.core.paths import to_windows_path
from m365_cli.core.text import clean_body, one_line, shorten_body, truncate, with_signoff
from m365_cli.core.timeutil import fmt, from_com, now, to_com
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
from m365_cli.features.mail.query import build_filter
from m365_cli.features.mail.rules import (
    ALL_FOLDERS,
    DigestRow,
    check_compose,
    group_digest,
    is_answered,
    pick_attachment,
    update_summary,
)

_BODY_START = re.compile(r"<body[^>]*>", re.I)
_PREVIEW_CHARS = 160
_THREAD_BODY_CHARS = 2000
_FLAGS = {"set": OlFlagStatus.MARKED, "clear": OlFlagStatus.NONE, "complete": OlFlagStatus.COMPLETE}
_HTML_FORMAT = 2
_DATA_IMAGE = re.compile(r"data:image/(png|jpe?g|gif);base64,([A-Za-z0-9+/=\s]+)", re.I)
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_MEETING_MESSAGE_CLASS = "IPM.Schedule"


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

    def _account(self) -> str:
        """SMTP address of the default account (to find its Outlook signature)."""
        return self._session.run(lambda ol: ol.smtp_of(ol.ns.CurrentUser.AddressEntry))

    @staticmethod
    def _sources(ol: Outlook, query: MailQuery) -> list[tuple[str, Com]]:
        if query.folder.strip().lower() == ALL_FOLDERS:
            return ol.mail_folders(query.mailbox)
        target = ol.folder(query.folder, query.mailbox)
        return [(target.Name, target)]

    @staticmethod
    def _matching(ol: Outlook, folder: Com, query: MailQuery) -> tuple[Com, bool]:
        """Items of `folder` matching the query, newest first; also whether it is the Sent folder."""
        items = folder.Items
        if dasl := build_filter(query):
            items = items.Restrict("@SQL=" + dasl)
        is_sent = folder.EntryID == ol.default_folder(OlFolder.SENT, query.mailbox).EntryID
        items.Sort("[SentOn]" if is_sent else "[ReceivedTime]", True)
        return items, is_sent

    def _newest(self, ol: Outlook, query: MailQuery, count: int) -> Iterator[tuple[str, Com, bool]]:
        """Up to `count` newest matching items over the query's folder(s) as (folder label, item, is_sent)."""
        streams = []
        for label, folder in self._sources(ol, query):
            items, is_sent = self._matching(ol, folder, query)
            field = "SentOn" if is_sent else "ReceivedTime"
            candidates = iterate(items)
            if query.has_attachments is not None:
                candidates = (it for it in candidates if _has_real_attachments(it) == query.has_attachments)
            head = [
                (from_com(get(it, field)) or datetime.min, label, it, is_sent)
                for it in itertools.islice(candidates, count)
            ]
            streams.append(head)
        merged = heapq.merge(*streams, key=lambda row: row[0], reverse=True)
        for _, label, item, is_sent in itertools.islice(merged, count):
            yield label, item, is_sent

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first. Folder "all" searches every mail folder."""
        everywhere = query.folder.strip().lower() == ALL_FOLDERS

        def run(ol: Outlook) -> Page[MailSummary]:
            rows = self._newest(ol, query, offset + query.limit + 1)
            found, more = take_page(rows, offset=offset, limit=query.limit)
            summaries = [
                self._summary(
                    ol,
                    it,
                    is_sent=sent,
                    preview=preview,
                    folder=label if everywhere else "",
                )
                for label, it, sent in found
            ]
            return Page(items=summaries, offset=offset, has_more=more)

        return self._session.run(run)

    def _summary(
        self,
        ol: Outlook,
        item: Com,
        *,
        is_sent: bool = False,
        preview: bool = False,
        folder: str = "",
    ) -> MailSummary:
        who = str(get(item, "To", "") if is_sent else get(item, "SenderName", ""))
        return MailSummary(
            ref=self._refs.put(ol.backend_id(item, KIND)),
            received=from_com(get(item, "SentOn" if is_sent else "ReceivedTime")),
            sender=("to " + who) if is_sent else who,
            subject=str(get(item, "Subject", "") or "(no subject)"),
            unread=bool(get(item, "UnRead", False)),
            has_attachments=_has_real_attachments(item),
            flagged=get(item, "FlagStatus", 0) == OlFlagStatus.MARKED,
            kind=kind_of(item),
            preview=one_line(clean_body(str(get(item, "Body", "") or "")), _PREVIEW_CHARS) if preview else "",
            folder=folder,
            protected=is_protected(item),
        )

    def select(self, query: MailQuery, *, max_items: int) -> list[str]:
        """Refs of up to `max_items` mails matching the query (cheap; for bulk updates)."""

        def run(ol: Outlook) -> list[str]:
            return [
                self._refs.put(ol.backend_id(item, KIND)) for _, item, _ in self._newest(ol, query, max_items)
            ]

        return self._session.run(run)

    def digest(self, query: MailQuery, *, max_items: int = 300) -> list[DigestGroup]:
        """Mails matching the query, grouped by sender (at most `max_items` mails are counted)."""

        def run(ol: Outlook) -> list[DigestGroup]:
            rows = [
                DigestRow(
                    sender=str(get(item, "SenderName", "") or ""),
                    subject=str(get(item, "Subject", "") or ""),
                    unread=bool(get(item, "UnRead", False)),
                    received=from_com(get(item, "ReceivedTime")),
                    ref=lambda it=item: self._refs.put(ol.backend_id(it, KIND)),
                )
                for _, item, _ in self._newest(ol, query, max_items)
            ]
            return group_digest(rows)

        return self._session.run(run)

    def _message(self, ol: Outlook, item: Com, *, full: bool, limit: int) -> MailMessage:
        body, quoted_cut, truncated = shorten_body(
            clean_body(str(get(item, "Body", "") or "")), full=full, limit=limit
        )
        return MailMessage(
            **self._summary(ol, item).model_dump(),
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

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        bid = self._refs.get(ref, kinds=KIND)

        def run(ol: Outlook) -> MailMessage:
            return self._message(ol, ol.item(bid), full=full, limit=self._settings.body_max_chars)

        return self._session.run(run)

    def mailboxes(self) -> list[tuple[str, str]]:
        """Mailboxes in the Outlook profile as (name, kind): primary | account | shared | other."""
        return self._session.run(lambda ol: ol.mailboxes())

    def thread(self, ref: str, *, full: bool = False) -> list[MailMessage]:
        """The whole conversation of a mail across folders, oldest first.

        Uses Outlook's conversation view; each message has its quoted history removed unless `full`.
        """
        bid = self._refs.get(ref, kinds=KIND)

        def run(ol: Outlook) -> list[MailMessage]:
            item = ol.item(bid)
            members = [item]
            conversation = get(item, "GetConversation")
            conv = conversation() if conversation else None
            if conv is not None:
                table = conv.GetTable()
                ids = []
                while not table.EndOfTable:
                    ids.append(table.GetNextRow()("EntryID"))
                members = [_item_or_none(ol, entry_id) for entry_id in ids] or [item]
            messages = [
                self._message(ol, m, full=full, limit=_THREAD_BODY_CHARS) for m in members if m is not None
            ]
            return sorted(messages, key=lambda m: m.received or datetime.min)

        return self._session.run(run)

    def waiting(
        self, *, days: int = 3, within_days: int = 14, limit: int = 20, mailbox: str | None = None
    ) -> list[WaitingItem]:
        """Sent mails from the last `within_days` days, older than `days`, with no reply yet.

        Meeting responses and mails to yourself are skipped. A reply is any later mail in the same
        conversation from someone else.
        """
        if days < 0 or within_days < days:
            raise InvalidInputError("Need 0 <= days <= within_days.")
        newest_allowed, oldest_allowed = now() - timedelta(days=days), now() - timedelta(days=within_days)

        def run(ol: Outlook) -> list[WaitingItem]:
            user = ol.ns.CurrentUser
            me = {str(user.Name).lower(), ol.smtp_of(user.AddressEntry).lower()}
            if mailbox:
                me.add(mailbox.lower())
            items = ol.default_folder(OlFolder.SENT, mailbox).Items
            items.Sort("[SentOn]", True)
            result: list[WaitingItem] = []
            for item in iterate(items):
                sent = from_com(get(item, "SentOn"))
                if sent is None or sent > newest_allowed:
                    continue
                if sent < oldest_allowed or len(result) >= limit:
                    break
                to = str(get(item, "To", "") or "")
                if str(get(item, "MessageClass", "")).startswith(_MEETING_MESSAGE_CLASS) or to.lower() in me:
                    continue
                if not is_answered(sent, str(item.Subject), me, _conversation_rows(item)):
                    result.append(
                        WaitingItem(
                            ref=self._refs.put(ol.backend_id(item, KIND)),
                            sent=sent,
                            to=to,
                            subject=str(item.Subject),
                            days=(now() - sent).days,
                        )
                    )
            return result

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
        body_format: BodyFormat = "markdown",
        attachments: list[str] | None = None,
        from_mailbox: str | None = None,
    ) -> Draft:
        """Save a draft (in Drafts) and return its preview. Never sends.

        `from_mailbox` sends from another account of the profile, or on behalf of a shared mailbox.
        A reply to a mail in another mailbox is sent from that mailbox automatically.

        The body is converted to HTML in the user's Outlook font, and Outlook's default signature
        (or the one set in the settings) is added, the same way Outlook does for a mail you write.

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

        kind: SignatureKind = "new" if source is None else "reply"
        sender = from_mailbox or (
            self._session.run(lambda ol: ol.sender_for_store(source.container_id)) if source else None
        )
        accounts = self._session.run(lambda ol: {str(a.SmtpAddress).lower() for a in ol.ns.Accounts})
        account = sender if sender and sender.lower() in accounts else self._account()
        signature = signature_for(kind, account, getattr(self._settings, f"signature_{kind}"))
        style = mail_style("compose" if source is None else "reply")
        fragment = styled(to_html(text, body_format), style) + (f"<br>{signature.html}" if signature else "")

        def run(ol: Outlook) -> Draft:
            if source is None:
                item = ol.app.CreateItem(OlItemType.MAIL)
                item.HTMLBody = html_document(_embed_images(item, fragment))
            else:
                original = ol.item(source)
                actions = {
                    "reply": original.Reply,
                    "reply_all": original.ReplyAll,
                    "forward": original.Forward,
                }
                item = actions[mode]()
                _insert_above_original(
                    item, _embed_images(item, fragment), to_text(text, body_format), signature
                )
            sent_as = ol.set_sender(item, sender) if sender else ""
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
                item.Close(OL_DISCARD)
                raise
            item.Save()
            return Draft(
                ref=self._refs.put(ol.backend_id(item, KIND)),
                mode=mode,
                to=_labels(ol, item, OlRecipientType.TO),
                cc=_labels(ol, item, OlRecipientType.CC),
                bcc=_labels(ol, item, OlRecipientType.BCC),
                subject=str(item.Subject),
                text=text,
                body_format=body_format,
                signature=signature.name if signature else "",
                from_mailbox=sent_as,
                includes_original=source is not None,
                attachments=[a.FileName for a in item.Attachments if not _is_inline(a)],
            )

        return self._session.run(run)

    def send(self, ref: str, *, send_at: datetime | None = None) -> str:
        """Send a draft now or, with `send_at`, schedule it (rate-limited and logged).

        A scheduled mail waits in the Outbox; Exchange delivers it at `send_at`.

        Raises:
            InvalidInputError: The item is not a draft, or `send_at` is in the past.
            RateLimitError: Too many sends in the last minute.
        """
        if send_at is not None and send_at <= now():
            raise InvalidInputError(f"send_at {fmt(send_at)} is in the past.")
        bid = self._refs.get(ref, kinds=KIND)
        self._audit.guard("mail_send")

        def run(ol: Outlook) -> tuple[str, str]:
            item = ol.item(bid)
            if get(item, "Sent", False):
                raise InvalidInputError(f"{ref} is not a draft (already sent/received).")
            ol.resolve_all(item)
            if send_at is not None:
                item.DeferredDeliveryTime = to_com(send_at)
            subject, recipients = str(item.Subject), "; ".join(r.Name for r in item.Recipients)
            item.Send()
            return subject, recipients

        subject, recipients = self._session.run(run)
        when = f" (scheduled for {fmt(send_at)})" if send_at else ""
        self._audit.record("mail_send", f"'{subject}' to {recipients}{when}")
        return f"{'Scheduled' if send_at else 'Sent'} '{subject}' to {recipients}{when}."

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
        bids = [self._refs.get(ref, kinds=KIND) for ref in refs]

        def run(ol: Outlook) -> str:
            target = ol.folder(move_to) if move_to else None
            moved_refs = []
            for bid in bids:
                item = ol.item(bid)
                if read is not None:
                    item.UnRead = not read
                if flag is not None:
                    item.FlagStatus = _FLAGS[flag]
                if categories is not None:
                    item.Categories = ", ".join(categories)
                if read is not None or flag is not None or categories is not None:
                    item.Save()
                if delete:
                    item.Delete()
                elif target is not None:
                    moved = item.Move(target)
                    moved_refs.append(self._refs.put(ol.backend_id(moved, KIND, target.StoreID)))
            return update_summary(
                len(bids),
                read=read,
                flag=flag,
                categories=categories,
                moved_to=target.Name if target is not None else None,
                moved_refs=moved_refs,
                delete=delete,
            )

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

    def attachment_text(self, ref: str, which: str) -> AttachmentText:
        """Extracted text of one attachment, chosen by 1-based index or by file name."""
        bid = self._refs.get(ref, kinds=KIND)
        with tempfile.TemporaryDirectory(prefix="m365-att-") as tmp:

            def run(ol: Outlook) -> Path:
                item = ol.item(bid)
                chosen = pick_attachment([(a.index, a.name) for a in attachments_of(item)], which)
                path = Path(tmp) / chosen[1]
                item.Attachments.Item(chosen[0]).SaveAsFile(str(path))
                return path

            path = self._session.run(run)
            text, cut = truncate(clean_body(extract_text(path)), self._settings.body_max_chars)
            return AttachmentText(name=path.name, text=text, truncated=cut)


def _has_real_attachments(item: Com) -> bool:
    """True when the item has attachments other than inline images (Outlook's flag counts those too)."""
    return any(not a.inline for a in attachments_of(item))


def _item_or_none(ol: Outlook, entry_id: str) -> Com | None:
    try:
        return ol.ns.GetItemFromID(entry_id)
    except pywintypes.com_error:
        return None


def _conversation_rows(item: Com) -> list[tuple[str, str, datetime]]:
    """(sender, subject, received) of all messages in the item's conversation (empty if unavailable)."""
    conversation = get(item, "GetConversation")
    conv = conversation() if conversation else None
    if conv is None:
        return []
    table = conv.GetTable()
    for column in ("SenderName", "ReceivedTime"):
        table.Columns.Add(column)
    rows = []
    while not table.EndOfTable:
        row = table.GetNextRow()
        received = from_com(row("ReceivedTime"))
        if received is not None:
            rows.append((str(row("SenderName") or ""), str(row("Subject") or ""), received))
    return rows


def _labels(ol: Outlook, item: Com, rtype: OlRecipientType) -> list[str]:
    labels = []
    for r in item.Recipients:
        if r.Type == rtype:
            smtp = ol.smtp_of(r.AddressEntry) if r.Resolved else ""
            labels.append(f"{r.Name} <{smtp}>" if smtp and smtp.lower() != r.Name.lower() else r.Name)
    return labels


def _insert_above_original(item: Com, fragment: str, text: str, signature: Signature | None) -> None:
    """Insert our HTML above the quoted original (plain text for plain-text originals)."""
    if item.BodyFormat == _HTML_FORMAT:
        body = item.HTMLBody
        m = _BODY_START.search(body)
        at = m.end() if m else 0
        item.HTMLBody = f"{body[:at]}{fragment}<br>{body[at:]}"
    else:
        sig = f"\n\n{signature.text}" if signature else ""
        item.Body = f"{text}{sig}\n\n{item.Body}"


def _embed_images(item: Com, fragment: str) -> str:
    """Turn data-URI images (e.g. a signature logo) into inline attachments referenced by cid:.

    Outlook does not send data: URIs reliably, so recipients would see a broken image otherwise.
    """
    images = list(_DATA_IMAGE.finditer(fragment))
    if not images:
        return fragment
    folder = Path(tempfile.mkdtemp(prefix="m365-img-"))
    for n, match in enumerate(images, start=1):
        cid = f"m365img{n}@m365-cli"
        path = folder / f"image{n:03}.{match.group(1)}"
        path.write_bytes(base64.b64decode(match.group(2)))
        attachment = item.Attachments.Add(str(path))
        attachment.PropertyAccessor.SetProperty(_PR_ATTACH_CONTENT_ID, cid)
        fragment = fragment.replace(match.group(0), f"cid:{cid}", 1)
    return fragment


def _is_inline(attachment: Com) -> bool:
    try:
        return bool(attachment.PropertyAccessor.GetProperty(_PR_ATTACH_CONTENT_ID))
    except pywintypes.com_error:
        return False
