from __future__ import annotations

import base64
import re
import tempfile
import urllib.parse
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from m365_cli.backends.outlook_web.session import ews_id, from_utc, quote, to_utc
from m365_cli.core.context import AppContext
from m365_cli.core.doctext import extract_text
from m365_cli.core.errors import InvalidInputError, NotFoundError, UnresolvedRecipientsError
from m365_cli.core.markup import (
    BodyFormat,
    MailStyle,
    Signature,
    body_fragment,
    html_document,
    parse_style,
    styled,
    to_html,
)
from m365_cli.core.paging import Page, take_page
from m365_cli.core.paths import local_path
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
    pick_attachment,
    update_summary,
)

_CONTENT_CLASS = "String {00020386-0000-0000-C000-000000000046} Name content-class"
_MESSAGE_CLASS = "String 0x001A"
_DEFERRED_SEND = "SystemTime 0x3FEF"
_PROTECTED_CLASS = "rpmsg.message"
_EXPAND_CLASS = (
    "SingleValueExtendedProperties($filter="
    f"PropertyId eq '{_MESSAGE_CLASS}' or PropertyId eq '{_CONTENT_CLASS}')"
)
_EXPAND_ATTACHMENTS = "Attachments($select=Name,Size,IsInline)"
_SUMMARY = (
    "Id,Subject,ReceivedDateTime,SentDateTime,From,ToRecipients,IsRead,HasAttachments,Flag,"
    "ParentFolderId,IsDraft"
)
_MESSAGE = f"{_SUMMARY},CcRecipients,BccRecipients,Categories,Body,ConversationId"
_WELL_KNOWN = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sent items": "sentitems",
    "sentitems": "sentitems",
    "drafts": "drafts",
    "deleted": "deleteditems",
    "deleted items": "deleteditems",
    "deleteditems": "deleteditems",
    "junk": "junkemail",
    "junk email": "junkemail",
    "junkemail": "junkemail",
    "archive": "archive",
    "outbox": "outbox",
}
_FLAGS = {"set": "Flagged", "clear": "NotFlagged", "complete": "Complete"}
_PREVIEW_CHARS = 160
_THREAD_BODY_CHARS = 2000
_SEARCH_PAGE = 250
_SEARCH_MAX = 1000
_SIMPLE_ATTACHMENT = 3 * 1024 * 1024
_MAX_ATTACHMENT = 150 * 1024 * 1024
_BODY_START = re.compile(r"<body[^>]*>", re.I)
_DATA_IMAGE = re.compile(r"data:image/(png|jpe?g|gif);base64,([A-Za-z0-9+/=\s]+)", re.I)
_MEETING_TYPES = ("EventMessage", "EventMessageRequest", "EventMessageResponse")
_SAFE = "=',()$:/"
_SKIPPED_FOLDERS = frozenset(
    {"deleteditems", "junkemail", "outbox", "drafts", "conversationhistory", "syncissues", "searchfolders"}
)
_SYSTEM_FOLDERS = frozenset(
    {"conversation action settings", "conversation history", "rss feeds", "sync issues", "yammer root"}
)


class WebMailService:
    """Mail operations on Outlook on the web. Every public method is safe to call from any thread.

    Refs hold the immutable item id and the mailbox ("" for your own), so they stay valid when a
    mail is moved or a draft is sent. Sending is two-step, as with classic Outlook: `compose()` saves
    a draft and returns a preview; `send()` sends it.
    """

    def __init__(self, ctx: AppContext) -> None:
        """Use the Outlook web session, refs, send log and settings of `ctx`."""
        self._web = ctx.outlook_web
        self._refs = ctx.refs
        self._audit = ctx.audit
        self._settings = ctx.settings

    def _ref(self, item_id: str, mailbox: str | None) -> str:
        return self._refs.put(BackendId(KIND, item_id, mailbox or ""))

    def _item(self, ref: str) -> tuple[str, str | None]:
        bid = self._refs.get(ref, kinds=KIND)
        return bid.item_id, bid.container_id or None

    def _folder(self, name: str, mailbox: str | None) -> tuple[str, str]:
        """(id, path) of a folder: a well-known name or a path such as 'Inbox/Projects'.

        Raises:
            NotFoundError: No such folder.
        """
        parts = [p for p in name.strip().strip("/").split("/") if p]
        if not parts:
            raise InvalidInputError("Empty folder name.")
        folders = self._web.folders(mailbox)
        well_known = _WELL_KNOWN.get(parts[0].lower())
        first = next((f for f in folders if well_known and f["WellKnownName"] == well_known), None)
        first = first or next((f for f in folders if f["path"].lower() == parts[0].lower()), None)
        if first is None:
            raise NotFoundError(f"Folder {parts[0]!r} not found.")
        path = "/".join([first["path"], *parts[1:]]).lower()
        found = next((f for f in folders if f["path"].lower() == path), None)
        if found is None:
            raise NotFoundError(f"Folder {name!r} not found.")
        return str(found["Id"]), str(found["path"])

    def _folder_path(self, folder_id: str, mailbox: str | None) -> str:
        return next((f["path"] for f in self._web.folders(mailbox) if f["Id"] == folder_id), "")

    def _searchable(self, mailbox: str | None) -> set[str]:
        """Ids of the folders that 'all' covers: not Deleted Items, Junk, Outbox, Drafts or system folders."""
        folders = self._web.folders(mailbox)
        skipped = [
            f["path"].lower() + "/"
            for f in folders
            if f["WellKnownName"] in _SKIPPED_FOLDERS
            or str(f.get("DisplayName", "")).lower() in _SYSTEM_FOLDERS
        ]
        return {f["Id"] for f in folders if not any((f["path"].lower() + "/").startswith(p) for p in skipped)}

    def _messages(
        self, query: MailQuery, count: int, *, offset: int = 0, select: str = _SUMMARY
    ) -> tuple[list[dict[str, Any]], bool, str | None]:
        """Matching messages, newest first: (`count` items after `offset`, more follow, mailbox)."""
        mailbox = self._web.mailbox(query.mailbox)
        everywhere = query.folder.strip().lower() == ALL_FOLDERS
        base = (
            "/messages" if everywhere else f"/mailfolders/{self._folder(query.folder, mailbox)[0]}/messages"
        )
        if "ParentFolderId" not in select.split(","):
            select += ",ParentFolderId"
        allowed = self._searchable(mailbox) if everywhere else None
        if query.text:
            return (*self._search(base, query, count, offset, select, mailbox, allowed), mailbox)
        params = {"$select": select, "$orderby": "ReceivedDateTime desc", "$expand": _EXPAND_CLASS}
        if flt := _filter(query):
            params["$filter"] = flt
        if allowed is None:
            params |= {"$top": str(count + 1), "$skip": str(offset)}
            items = (self._web.rest(f"{base}?{_params(params)}", mailbox=mailbox) or {}).get("value", [])
            return items[:count], len(items) > count, mailbox
        params["$top"] = str(min(250, 2 * (offset + count + 1)))
        found = self._web.rest_pages(f"{base}?{_params(params)}", max_items=_SEARCH_MAX, mailbox=mailbox)
        items, more = take_page(
            (m for m in found if m.get("ParentFolderId") in allowed), offset=offset, limit=count
        )
        return items, more, mailbox

    def _search(
        self,
        base: str,
        query: MailQuery,
        count: int,
        offset: int,
        select: str,
        mailbox: str | None,
        allowed: set[str] | None,
    ) -> tuple[list[dict[str, Any]], bool]:
        """Full-text search (the server allows no other filter or order with it), then the other filters."""
        text = (query.text or "").replace('"', " ").strip()
        params = {
            "$search": f'"{text}"',
            "$select": select,
            "$top": str(_SEARCH_PAGE),
            "$expand": _EXPAND_CLASS,
        }
        found = self._web.rest_pages(f"{base}?{_params(params)}", max_items=_SEARCH_MAX, mailbox=mailbox)
        matching = (
            m for m in found if _matches(m, query) and (allowed is None or m.get("ParentFolderId") in allowed)
        )
        items, more = take_page(matching, offset=offset, limit=count)
        ids = self._web.immutable_ids([m["Id"] for m in items], mailbox=mailbox)
        return [{**m, "Id": new} for m, new in zip(items, ids, strict=True)], more

    def _summary(
        self,
        m: dict[str, Any],
        mailbox: str | None,
        *,
        folder: str = "",
        preview: bool = False,
        as_list: bool = True,
    ) -> MailSummary:
        """List entry; in a list, sent mails show their recipients and sent time (not when reading)."""
        sent = as_list and self._is_sent(m, mailbox)
        who = (
            "; ".join(_name(r) for r in m.get("ToRecipients") or [])
            if sent
            else _name(self._from(m, mailbox))
        )
        return MailSummary(
            ref=self._ref(m["Id"], mailbox),
            received=from_utc(m.get("SentDateTime") if sent else m.get("ReceivedDateTime")),
            sender=("to " + who) if sent else who,
            subject=str(m.get("Subject") or "(no subject)"),
            unread=not m.get("IsRead", True),
            has_attachments=bool(m.get("HasAttachments")),
            flagged=(m.get("Flag") or {}).get("FlagStatus") == "Flagged",
            kind=_kind(m),
            preview=one_line(clean_body(str(m.get("BodyPreview") or "")), _PREVIEW_CHARS) if preview else "",
            folder=folder,
            protected=_is_protected(m),
        )

    def _from(self, m: dict[str, Any], mailbox: str | None) -> dict[str, Any] | None:
        """The sender; a draft without one goes out from its mailbox (yours, or the shared one)."""
        if m.get("From") or not m.get("IsDraft"):
            return m.get("From")
        name, address = self._web.me()
        if mailbox:
            name = next((n for n, a in self._web.shared_mailboxes() if a == mailbox), mailbox)
        return {"EmailAddress": {"Name": name, "Address": mailbox or address}}

    def _is_sent(self, m: dict[str, Any], mailbox: str | None) -> bool:
        return m.get("ParentFolderId") == self._folder("sent", mailbox)[0]

    def find(self, query: MailQuery, *, offset: int = 0, preview: bool = False) -> Page[MailSummary]:
        """One page of mails matching the query, newest first. Folder "all" searches every mail folder.

        Text uses the server's full-text search over all fields; the other filters are exact.
        """
        everywhere = query.folder.strip().lower() == ALL_FOLDERS
        select = f"{_SUMMARY},BodyPreview" if preview else _SUMMARY
        items, more, mailbox = self._messages(query, query.limit, offset=offset, select=select)
        summaries = [
            self._summary(
                m,
                mailbox,
                folder=self._folder_path(str(m.get("ParentFolderId")), mailbox) if everywhere else "",
                preview=preview,
            )
            for m in items
        ]
        return Page(items=summaries, offset=offset, has_more=more)

    def select(self, query: MailQuery, *, max_items: int) -> list[str]:
        """Refs of up to `max_items` mails matching the query (cheap; for bulk updates)."""
        items, _, mailbox = self._messages(query, max_items, select="Id")
        return [self._ref(m["Id"], mailbox) for m in items]

    def digest(self, query: MailQuery, *, max_items: int = 300) -> list[DigestGroup]:
        """Mails matching the query, grouped by sender (at most `max_items` mails are counted)."""
        items, _, mailbox = self._messages(query, max_items, select="Id,Subject,From,IsRead,ReceivedDateTime")
        rows = [
            DigestRow(
                sender=_name(m.get("From")),
                subject=str(m.get("Subject") or ""),
                unread=not m.get("IsRead", True),
                received=from_utc(m.get("ReceivedDateTime")),
                ref=lambda m=m: self._ref(m["Id"], mailbox),
            )
            for m in items
        ]
        return group_digest(rows)

    def mailboxes(self) -> list[tuple[str, str]]:
        """Your mailbox and the other mailboxes added to Outlook on the web, as (name, kind)."""
        name, address = self._web.me()
        return [(f"{name} <{address}>", "primary")] + [
            (f"{n} <{a}>", "shared") for n, a in self._web.shared_mailboxes()
        ]

    def _get(self, item_id: str, mailbox: str | None, select: str = _MESSAGE) -> dict[str, Any]:
        path = f"/messages/{item_id}?$select={select}&$expand={_EXPAND_CLASS},{_EXPAND_ATTACHMENTS}"
        return self._web.rest(path, mailbox=mailbox, text=True) or {}

    def _message(self, m: dict[str, Any], mailbox: str | None, *, full: bool, limit: int) -> MailMessage:
        summary = self._summary(m, mailbox, as_list=False)
        text = str((m.get("Body") or {}).get("Content") or "")
        if summary.protected and self._settings.read_protected:
            decrypted = self._decrypted(m["Id"], mailbox)
            if decrypted is not None:
                text = decrypted
                summary = summary.model_copy(update={"protected": False})
        body, quoted_cut, truncated = shorten_body(clean_body(text), full=full, limit=limit)
        return MailMessage(
            **summary.model_dump(),
            sender_email=_address(self._from(m, mailbox)),
            to="; ".join(_name(r) for r in m.get("ToRecipients") or []),
            cc="; ".join(_name(r) for r in m.get("CcRecipients") or []),
            categories=", ".join(m.get("Categories") or []),
            attachments=[
                Attachment(
                    index=i,
                    name=str(a.get("Name") or ""),
                    size=int(a.get("Size") or 0),
                    inline=bool(a.get("IsInline")),
                )
                for i, a in enumerate(m.get("Attachments") or [], start=1)
            ],
            body=body,
            quoted_cut=quoted_cut,
            truncated=truncated,
        )

    def _decrypted(self, item_id: str, mailbox: str | None) -> str | None:
        """Text of a rights-protected mail decrypted by Outlook on the web; None when it could not decrypt.

        Used only with read_protected. S/MIME mail and mail you have no rights to stay protected.
        """
        answer = self._web.svc(
            "GetItem",
            {
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "BodyType": "Text",
                    "ClientSupportsIrm": True,
                    "AdditionalProperties": [
                        {"__type": "PropertyUri:#Exchange", "FieldURI": "Body"},
                        {"__type": "PropertyUri:#Exchange", "FieldURI": "RightsManagementLicenseData"},
                    ],
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": ews_id(item_id)}],
            },
            mailbox=mailbox,
        )
        item = (answer.get("Items") or [{}])[0]
        license_data = item.get("RightsManagementLicenseData") or {}
        if license_data.get("RightsManagedMessageDecryptionStatus") != 0:
            return None
        return str((item.get("Body") or {}).get("Value") or "")

    def read(self, ref: str, *, full: bool = False) -> MailMessage:
        """One mail; unless `full`, quoted history is omitted and long bodies are cut."""
        item_id, mailbox = self._item(ref)
        return self._message(
            self._get(item_id, mailbox), mailbox, full=full, limit=self._settings.body_max_chars
        )

    def thread(self, ref: str, *, full: bool = False) -> list[MailMessage]:
        """The whole conversation of a mail across folders, oldest first."""
        item_id, mailbox = self._item(ref)
        conversation = (
            self._web.rest(f"/messages/{item_id}?$select=ConversationId", mailbox=mailbox) or {}
        ).get("ConversationId")
        if not conversation:
            return [self._message(self._get(item_id, mailbox), mailbox, full=full, limit=_THREAD_BODY_CHARS)]
        params = {
            "$filter": f"ConversationId eq '{quote(str(conversation))}'",
            "$select": _MESSAGE,
            "$top": "100",
            "$expand": f"{_EXPAND_CLASS},{_EXPAND_ATTACHMENTS}",
        }
        members = self._web.rest_pages(
            f"/messages?{_params(params)}", max_items=200, mailbox=mailbox, text=True
        )
        messages = [self._message(m, mailbox, full=full, limit=_THREAD_BODY_CHARS) for m in members]
        return sorted(messages, key=lambda m: m.received or datetime.min)

    def waiting(
        self, *, days: int = 3, within_days: int = 14, limit: int = 20, mailbox: str | None = None
    ) -> list[WaitingItem]:
        """Sent mails from the last `within_days` days, older than `days`, with no reply yet.

        Meeting messages and mails to yourself are skipped. A reply is a later mail from someone else
        in the same conversation.
        """
        if days < 0 or within_days < days:
            raise InvalidInputError("Need 0 <= days <= within_days.")
        box = self._web.mailbox(mailbox)
        newest_allowed, oldest_allowed = now() - timedelta(days=days), now() - timedelta(days=within_days)
        name, address = self._web.me()
        me = {name.lower(), address.lower()} | ({box} if box else set())
        sent_id = self._folder("sent", box)[0]
        sent = self._web.rest_pages(
            f"/mailfolders/{sent_id}/messages?"
            + _params(
                {
                    "$filter": f"SentDateTime ge {to_utc(oldest_allowed)} "
                    f"and SentDateTime le {to_utc(newest_allowed)}",
                    "$orderby": "SentDateTime desc",
                    "$select": "Id,Subject,SentDateTime,ToRecipients,ConversationId",
                    "$top": "200",
                    "$expand": _EXPAND_CLASS,
                }
            ),
            max_items=1000,
            mailbox=box,
        )
        later = self._web.rest_pages(
            "/messages?"
            + _params(
                {
                    "$filter": f"ReceivedDateTime ge {to_utc(oldest_allowed)}",
                    "$select": "ConversationId,From,Subject,ReceivedDateTime",
                    "$top": "500",
                }
            ),
            max_items=5000,
            mailbox=box,
        )
        by_conversation: dict[str, list[tuple[str, str, datetime]]] = defaultdict(list)
        for m in later:
            received = from_utc(m.get("ReceivedDateTime"))
            if received is not None:
                by_conversation[str(m.get("ConversationId"))].append(
                    (_name(m.get("From")) or _address(m.get("From")), str(m.get("Subject") or ""), received)
                )
        result: list[WaitingItem] = []
        for m in sent:
            when = from_utc(m.get("SentDateTime"))
            recipients = m.get("ToRecipients") or []
            to = "; ".join(_name(r) for r in recipients)
            to_self = all(_address(r).lower() in me or _name(r).lower() in me for r in recipients)
            if when is None or _kind(m) == "meeting" or to_self:
                continue
            rows = by_conversation.get(str(m.get("ConversationId")), [])
            if not is_answered(when, str(m.get("Subject") or ""), me, rows):
                result.append(
                    WaitingItem(
                        ref=self._ref(m["Id"], box),
                        sent=when,
                        to=to,
                        subject=str(m.get("Subject") or ""),
                        days=(now() - when).days,
                    )
                )
                if len(result) >= limit:
                    break
        return result

    def _recipients(self, values: list[str] | None) -> tuple[list[dict[str, Any]], list[str]]:
        """API recipients for names/addresses, plus the values that did not resolve."""
        resolved, missing = [], []
        for value in values or []:
            found = self._web.resolve(value)
            if found is None:
                missing.append(value)
            else:
                resolved.append({"EmailAddress": {"Name": found[0], "Address": found[1]}})
        return resolved, missing

    def _signature(self, kind: str) -> Signature | None:
        """Your roaming signature for new mail or replies: the setting's name, Outlook's default, or none."""
        setting = str(getattr(self._settings, f"signature_{kind}", "auto")).strip()
        if setting.lower() == "none":
            return None
        name = setting
        if setting.lower() == "auto":
            key = "roaming_new_signature" if kind == "new" else "roaming_reply_signature"
            values = {s.get("name"): str(s.get("value") or "") for s in self._web.cloud_settings(key)}
            name = values.get(key, "").strip()
            if not name:
                return None
        entries = [str(s.get("value") or "") for s in self._web.cloud_settings(name, account=True)]
        html = next((v for v in entries if "<" in v), "")
        text = next((v for v in entries if v and "<" not in v), "")
        if not html and not text:
            raise NotFoundError(f"Signature {name!r} not found in your roaming signatures.")
        return Signature(
            name=name, html=body_fragment(html) if html else text.replace("\n", "<br>"), text=text
        )

    def _style(self) -> MailStyle:
        if self._settings.mail_style:
            return parse_style(self._settings.mail_style)
        options = self._web.user_options()
        family = str(options.get("ComposeFontName") or "")
        size = str(options.get("ComposeFontSize") or "")
        color = str(options.get("ComposeFontColor") or "")
        return MailStyle(
            font_family="" if family in ("", "None") else family,
            font_size_pt=float(size) if size.replace(".", "", 1).isdigit() else None,
            color=color if color.startswith("#") else "",
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
        body_format: BodyFormat = "markdown",
        attachments: list[str] | None = None,
        from_mailbox: str | None = None,
    ) -> Draft:
        """Save a draft and return its preview. Never sends.

        A reply or forward to a mail in another mailbox is saved in that mailbox's Drafts and goes out
        from it; `from_mailbox` does the same for a new mail. The body is written in your font, with
        your roaming signature (for new mail and for replies).

        Raises:
            InvalidInputError: Arguments don't fit the mode, or an attachment file is missing or too big.
            UnresolvedRecipientsError: A recipient is unknown; nothing is saved.
        """
        check_compose(mode, ref=ref, to=to, subject=subject)
        files = _attachment_files(attachments)
        groups = {key: self._recipients(values) for key, values in (("To", to), ("Cc", cc), ("Bcc", bcc))}
        missing = [value for _, miss in groups.values() for value in miss]
        if missing:
            raise UnresolvedRecipientsError(missing)
        source = self._item(ref) if ref else None
        sender = self._web.mailbox(from_mailbox) if from_mailbox else None
        home = source[1] if source else sender
        text = with_signoff(body, self._settings.signoff)
        signature = self._signature("new" if source is None else "reply")
        fragment = styled(to_html(text, body_format), self._style()) + (
            f"<br>{signature.html}" if signature else ""
        )
        fragment, images = _inline_images(fragment)
        if source is None:
            draft = self._web.rest(
                "/messages",
                method="POST",
                mailbox=home,
                body={
                    "Subject": subject or "",
                    "Body": {"ContentType": "HTML", "Content": html_document(fragment)},
                    **{f"{k}Recipients": v[0] for k, v in groups.items() if v[0]},
                },
            )
        else:
            action = {"reply": "createreply", "reply_all": "createreplyall", "forward": "createforward"}[mode]
            draft = self._web.rest(f"/messages/{source[0]}/{action}", method="POST", mailbox=home, body={})
        draft_id = str(draft["Id"])
        try:
            if source is not None:
                change = _reply_change(draft, fragment, {k: v[0] for k, v in groups.items()}, subject)
                if sender and sender != home:
                    change["From"] = {"EmailAddress": {"Address": sender}}
                draft = self._web.rest(f"/messages/{draft_id}", method="PATCH", mailbox=home, body=change)
            for n, (cid, ext, data) in enumerate(images, start=1):
                self._attach(
                    draft_id, home, f"image{n:03}.{ext}", data, content_type=f"image/{ext}", content_id=cid
                )
            for f in files:
                self._attach_file(draft_id, home, f)
            if not any(draft.get(f"{k}Recipients") for k in ("To", "Cc", "Bcc")):
                raise InvalidInputError("The draft has no recipients.")
        except Exception:
            self._web.rest(f"/messages/{draft_id}", method="DELETE", mailbox=home)
            raise
        return Draft(
            ref=self._ref(draft_id, home),
            mode=mode,
            to=[_label(r) for r in draft.get("ToRecipients") or []],
            cc=[_label(r) for r in draft.get("CcRecipients") or []],
            bcc=[_label(r) for r in draft.get("BccRecipients") or []],
            subject=str(draft.get("Subject") or ""),
            text=text,
            body_format=body_format,
            signature=signature.name if signature else "",
            from_mailbox=sender or home or "",
            includes_original=source is not None,
            attachments=[f.name for f in files],
            saved_in=self._drafts_location(home),
        )

    def _drafts_location(self, mailbox: str | None) -> str:
        """Mailbox and folder a draft is saved in, such as 'Github / Drafts'."""
        name = self._web.me()[0]
        if mailbox:
            name = next((n for n, a in self._web.shared_mailboxes() if a == mailbox), mailbox)
        return f"{name} / {self._folder('drafts', mailbox)[1]}"

    def _attach_file(self, draft_id: str, mailbox: str | None, path: Path) -> None:
        """Attach a file: up to 3 MB in one request, bigger ones through an upload session."""
        data = path.read_bytes()
        if len(data) <= _SIMPLE_ATTACHMENT:
            self._attach(draft_id, mailbox, path.name, base64.b64encode(data).decode())
            return
        item = {"AttachmentType": "File", "Name": path.name, "Size": len(data)}
        session = self._web.rest(
            f"/messages/{draft_id}/attachments/createUploadSession",
            method="POST",
            mailbox=mailbox,
            body={"AttachmentItem": item},
        )
        self._web.upload(str((session or {})["UploadUrl"]), data)

    def _attach(
        self,
        draft_id: str,
        mailbox: str | None,
        name: str,
        data: str,
        *,
        content_type: str = "",
        content_id: str = "",
    ) -> None:
        body: dict[str, Any] = {
            "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
            "Name": name,
            "ContentBytes": data,
        }
        if content_type:
            body["ContentType"] = content_type
        if content_id:
            body |= {"IsInline": True, "ContentId": content_id}
        self._web.rest(f"/messages/{draft_id}/attachments", method="POST", mailbox=mailbox, body=body)

    def send(self, ref: str, *, send_at: datetime | None = None) -> str:
        """Send a draft now or, with `send_at`, schedule it (rate-limited and logged).

        A scheduled mail waits in Drafts; Exchange delivers it at `send_at`.

        Raises:
            InvalidInputError: The item is not a draft, is already scheduled, or `send_at` is in the past.
            RateLimitError: Too many sends in the last minute.
        """
        if send_at is not None and send_at <= now():
            raise InvalidInputError(f"send_at {fmt(send_at)} is in the past.")
        item_id, mailbox = self._item(ref)
        draft = (
            self._web.rest(
                f"/messages/{item_id}?$select=IsDraft,Subject,ToRecipients,CcRecipients,BccRecipients"
                f"&$expand=SingleValueExtendedProperties($filter=PropertyId eq '{_DEFERRED_SEND}')",
                mailbox=mailbox,
            )
            or {}
        )
        if not draft.get("IsDraft"):
            raise InvalidInputError(f"{ref} is not a draft (already sent/received).")
        scheduled = from_utc(_property(draft, _DEFERRED_SEND))
        if scheduled is not None and scheduled > now():
            raise InvalidInputError(f"{ref} is already scheduled for {fmt(scheduled)}.")
        self._audit.guard("mail_send")
        if send_at is not None:
            self._web.rest(
                f"/messages/{item_id}",
                method="PATCH",
                mailbox=mailbox,
                body={
                    "SingleValueExtendedProperties": [
                        {"PropertyId": _DEFERRED_SEND, "Value": to_utc(send_at)}
                    ]
                },
            )
        self._web.rest(f"/messages/{item_id}/send", method="POST", mailbox=mailbox)
        subject = str(draft.get("Subject") or "")
        recipients = "; ".join(
            _name(r) for k in ("To", "Cc", "Bcc") for r in draft.get(f"{k}Recipients") or []
        )
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
        """Mark read/unread, flag, set categories, move or delete one or more mails; returns a summary.

        Delete moves the mail to Deleted Items. Refs stay the same after a move.
        """
        if not refs:
            raise InvalidInputError("No mails to update.")
        if read is None and flag is None and categories is None and not move_to and not delete:
            raise InvalidInputError("Nothing to update.")
        items = [self._item(ref) for ref in refs]
        change: dict[str, Any] = {}
        if read is not None:
            change["IsRead"] = read
        if flag is not None:
            change["Flag"] = {"FlagStatus": _FLAGS[flag]}
        if categories is not None:
            change["Categories"] = categories
        moved_refs: list[str] = []
        target_name = None
        for item_id, mailbox in items:
            if change:
                self._web.rest(f"/messages/{item_id}", method="PATCH", mailbox=mailbox, body=change)
            if delete or move_to:
                target_id, target_name = self._folder("deleted" if delete else str(move_to), mailbox)
                moved = (
                    self._web.rest(
                        f"/messages/{item_id}/move",
                        method="POST",
                        mailbox=mailbox,
                        body={"DestinationId": target_id},
                    )
                    or {}
                )
                moved_refs.append(self._ref(str(moved.get("Id") or item_id), mailbox))
        return update_summary(
            len(items),
            read=read,
            flag=flag,
            categories=categories,
            moved_to=target_name if move_to and not delete else None,
            moved_refs=moved_refs,
            delete=delete,
        )

    def _attachment_list(self, item_id: str, mailbox: str | None) -> list[dict[str, Any]]:
        reply = (
            self._web.rest(f"/messages/{item_id}/attachments?$select=Name,Size,IsInline", mailbox=mailbox)
            or {}
        )
        return list(reply.get("value", []))

    def _download(self, item_id: str, mailbox: str | None, attachment: dict[str, Any]) -> tuple[str, bytes]:
        """(file name, bytes) of one attachment; an attached mail is saved as .eml."""
        name = _safe_name(str(attachment.get("Name") or "attachment"))
        path = f"/messages/{item_id}/attachments/{attachment['Id']}"
        if str(attachment.get("@odata.type", "")).endswith("ItemAttachment"):
            return f"{name}.eml", self._web.rest(f"{path}/$value", mailbox=mailbox, binary=True)
        full = self._web.rest(path, mailbox=mailbox) or {}
        return name, base64.b64decode(str(full.get("ContentBytes") or ""))

    def attachments(self, ref: str, *, save_to: str | None = None, include_inline: bool = False) -> list[str]:
        """List the attachments, or save them into directory `save_to`."""
        item_id, mailbox = self._item(ref)
        infos = [
            (i, a)
            for i, a in enumerate(self._attachment_list(item_id, mailbox), start=1)
            if include_inline or not a.get("IsInline")
        ]
        if save_to is None:
            return [f"{i}. {a.get('Name')} ({max(1, int(a.get('Size') or 0) // 1024)} KB)" for i, a in infos]
        target = Path(local_path(save_to))
        target.mkdir(parents=True, exist_ok=True)
        saved = []
        for _, a in infos:
            name, data = self._download(item_id, mailbox, a)
            path = target / name
            path.write_bytes(data)
            saved.append(str(path))
        return saved

    def attachment_text(self, ref: str, which: str) -> AttachmentText:
        """Extracted text of one attachment, chosen by 1-based index or by file name."""
        item_id, mailbox = self._item(ref)
        listing = self._attachment_list(item_id, mailbox)
        index, _ = pick_attachment(
            [(i, str(a.get("Name") or "")) for i, a in enumerate(listing, start=1)], which
        )
        name, data = self._download(item_id, mailbox, listing[index - 1])
        with tempfile.TemporaryDirectory(prefix="m365-att-") as tmp:
            path = Path(tmp) / name
            path.write_bytes(data)
            text, cut = truncate(clean_body(extract_text(path)), self._settings.body_max_chars)
        return AttachmentText(name=name, text=text, truncated=cut)


def _params(params: dict[str, str]) -> str:
    return "&".join(f"{k}={urllib.parse.quote(v, safe=_SAFE)}" for k, v in params.items())


def _filter(q: MailQuery) -> str:
    """OData filter for the query (without text). The sort field comes first, as the server requires."""
    parts = [f"ReceivedDateTime ge {to_utc(q.since) if q.since else '1900-01-01T00:00:00Z'}"]
    if q.until:
        parts.append(f"ReceivedDateTime lt {to_utc(q.until)}")
    if q.unread_only:
        parts.append("IsRead eq false")
    if q.has_attachments is not None:
        parts.append(f"HasAttachments eq {str(q.has_attachments).lower()}")
    if q.flagged is not None:
        parts.append("Flag/FlagStatus eq 'Flagged'" if q.flagged else "Flag/FlagStatus ne 'Flagged'")
    if q.subject:
        parts.append(f"contains(Subject,'{quote(q.subject)}')")
    if q.sender:
        s = quote(q.sender)
        parts.append(f"(contains(From/EmailAddress/Name,'{s}') or contains(From/EmailAddress/Address,'{s}'))")
    return " and ".join(parts) if len(parts) > 1 or q.since else ""


def _matches(m: dict[str, Any], q: MailQuery) -> bool:
    """The query's non-text conditions, checked on a search result."""
    received = from_utc(m.get("ReceivedDateTime"))
    checks = [
        not (q.since and (received is None or received < q.since)),
        not (q.until and (received is None or received >= q.until)),
        not (q.unread_only and m.get("IsRead", True)),
        q.has_attachments is None or bool(m.get("HasAttachments")) == q.has_attachments,
        q.flagged is None or ((m.get("Flag") or {}).get("FlagStatus") == "Flagged") == q.flagged,
        not q.subject or q.subject.lower() in str(m.get("Subject") or "").lower(),
        not q.sender or q.sender.lower() in f"{_name(m.get('From'))} {_address(m.get('From'))}".lower(),
    ]
    return all(checks)


def _property(m: dict[str, Any], property_id: str) -> str | None:
    """Value of an extended property; the server writes tags without leading zeros ('String 0x1a')."""
    key = _property_key(property_id)
    for p in m.get("SingleValueExtendedProperties") or []:
        if _property_key(str(p.get("PropertyId", ""))) == key:
            return str(p.get("Value"))
    return None


def _property_key(property_id: str) -> str:
    kind, _, tag = property_id.strip().partition(" ")
    if tag.lower().startswith("0x"):
        return f"{kind.lower()} {int(tag, 16)}"
    return property_id.strip().lower()


def _kind(m: dict[str, Any]) -> str:
    """Item kind: mail | meeting | report | other."""
    if str(m.get("@odata.type", "")).rsplit(".", 1)[-1] in _MEETING_TYPES:
        return "meeting"
    cls = (_property(m, _MESSAGE_CLASS) or "IPM.Note").lower()
    if cls.startswith("ipm.schedule"):
        return "meeting"
    if cls.startswith(("report.", "ipm.report")):
        return "report"
    return "mail" if cls.startswith("ipm.note") else "other"


def _is_protected(m: dict[str, Any]) -> bool:
    cls = (_property(m, _MESSAGE_CLASS) or "").lower()
    return _property(m, _CONTENT_CLASS) == _PROTECTED_CLASS or cls.startswith("ipm.note.smime")


def _name(recipient: dict[str, Any] | None) -> str:
    address = (recipient or {}).get("EmailAddress") or {}
    return str(address.get("Name") or address.get("Address") or "")


def _address(recipient: dict[str, Any] | None) -> str:
    return str(((recipient or {}).get("EmailAddress") or {}).get("Address") or "")


def _label(recipient: dict[str, Any]) -> str:
    name, address = _name(recipient), _address(recipient)
    return f"{name} <{address}>" if address and name.lower() != address.lower() else name or address


def _attachment_files(paths: list[str] | None) -> list[Path]:
    """Local attachment files, checked before anything is saved.

    Raises:
        InvalidInputError: A file is missing or bigger than the simple upload allows.
    """
    files = [Path(local_path(p)) for p in paths or []]
    for f in files:
        if not f.is_file():
            raise InvalidInputError(f"Attachment not found: {f}")
        if f.stat().st_size > _MAX_ATTACHMENT:
            raise InvalidInputError(f"Attachment {f.name} is larger than 150 MB; share a link instead.")
    return files


def _reply_change(
    draft: dict[str, Any], fragment: str, added: dict[str, list[dict[str, Any]]], subject: str | None
) -> dict[str, Any]:
    """Changes for a reply/forward draft: our text above the original, extra recipients, a new subject."""
    change: dict[str, Any] = {"Body": {"ContentType": "HTML", "Content": _above_original(draft, fragment)}}
    for key, people in added.items():
        if people:
            change[f"{key}Recipients"] = [*(draft.get(f"{key}Recipients") or []), *people]
    if subject is not None:
        change["Subject"] = subject
    return change


def _above_original(draft: dict[str, Any], fragment: str) -> str:
    """Our HTML above the quoted original of a reply or forward draft."""
    body = str((draft.get("Body") or {}).get("Content") or "")
    m = _BODY_START.search(body)
    if not m:
        return html_document(f"{fragment}<br>{body}")
    return f"{body[: m.end()]}{fragment}<br>{body[m.end() :]}"


def _inline_images(fragment: str) -> tuple[str, list[tuple[str, str, str]]]:
    """Data-URI images (e.g. a signature logo) as cid: references plus (content id, extension, base64)."""
    images: list[tuple[str, str, str]] = []
    for n, match in enumerate(_DATA_IMAGE.finditer(fragment), start=1):
        cid = f"m365img{n}@m365-cli"
        ext = match.group(1).lower().replace("jpg", "jpeg")
        images.append((cid, ext, re.sub(r"\s", "", match.group(2))))
        fragment = fragment.replace(match.group(0), f"cid:{cid}", 1)
    return fragment, images


def _safe_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "attachment"
