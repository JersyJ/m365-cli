from __future__ import annotations

import time
import urllib.parse
import uuid
from collections.abc import Iterator
from datetime import datetime
from functools import cached_property
from pathlib import Path
from typing import TYPE_CHECKING, Any

from m365_cli.core.doctext import extract_text
from m365_cli.core.errors import BackendUnavailableError, InvalidInputError, M365Error, NotFoundError
from m365_cli.core.markup import BodyFormat
from m365_cli.core.paging import Page, take_page
from m365_cli.core.refs import BackendId
from m365_cli.core.text import clean_body, one_line, truncate, with_signoff
from m365_cli.core.timeutil import now
from m365_cli.features.teams import rules
from m365_cli.features.teams.cache import ProfileCache
from m365_cli.features.teams.drafts import DraftFile
from m365_cli.features.teams.models import (
    Chat,
    ChatKind,
    Digest,
    FileText,
    Member,
    Mention,
    MessagePage,
    Presence,
    SearchHit,
    SharedFile,
    TeamsDraft,
    TeamsMessage,
)

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

CHAT = "chatServiceAfd"
MIDDLE_TIER = "middleTier"
SPACES = "https://api.spaces.skype.com"
SUBSTRATE = "substrate"
SUBSTRATE_AUD = "66a88757-258c-4c72-893c-3e8bed4d6899"
SUBSTRATE_OFFICE = "https://substrate.office.com"
PRESENCE = "presence"
SHAREPOINT_AUD = "00000003-0000-0ff1-ce00-000000000000"
_LIST_SIZE = 500
_MENTIONS = "48:mentions"
_FILE_SCAN = 40
_MESSAGE_VIEW = "msnp24Equivalent|supportsMessageProperties"
_PROFILE_PATH = "/beta/users/fetchShortProfile?isMailAddress={}&enableGuest=true&skypeTeamsInfo=true"


def _q(conversation_id: str) -> str:
    return urllib.parse.quote(conversation_id, safe=";=")


class TeamsService:
    """Teams operations against the real Teams web APIs (see `TeamsApi` for the contract)."""

    def __init__(self, ctx: AppContext) -> None:
        """Use the context's Teams session, refs, send log and settings."""
        self._ctx = ctx
        self._session = ctx.teams
        self._refs = ctx.refs
        self._drafts = DraftFile(ctx.settings.data_dir / "teams-drafts.json")
        self._profiles = ProfileCache(ctx.settings.data_dir / "teams-people.json")

    @cached_property
    def _me(self) -> Member:
        me = (self._session.status() or {}).get("me") or {}
        if not me.get("oid"):
            raise BackendUnavailableError("Teams has not loaded your account yet. Retry in a few seconds.")
        return Member(
            name=str(me.get("name") or ""), mri=f"8:orgid:{me['oid']}", email=str(me.get("upn") or "")
        )

    @cached_property
    def _conversations(self) -> list[dict[str, Any]]:
        path = f"/v1/users/ME/conversations?view=msnp24Equivalent&pageSize={_LIST_SIZE}"
        reply = self._session.request(CHAT, path) or {}
        now_ms = int(time.time() * 1000)
        return [c for c in reply.get("conversations") or [] if rules.is_listed(c, now_ms=now_ms)]

    def chats(
        self,
        *,
        query: str | None = None,
        kind: ChatKind | None = None,
        unread_only: bool = False,
        offset: int = 0,
        limit: int = 15,
    ) -> Page[Chat]:
        """See `TeamsApi.chats`."""
        found = self._all_chats()
        if query:
            found = [c for c in found if rules.matches(c.title, query)]
        if kind:
            found = [c for c in found if c.kind == kind]
        if unread_only:
            found = [c for c in found if c.unread]
        items, more = take_page(found, offset=offset, limit=limit)
        return Page(items=items, offset=offset, has_more=more)

    def _all_chats(self) -> list[Chat]:
        convs = self._conversations
        self._load_profiles(
            [m for c in convs if (m := rules.other_member_mri(str(c.get("id", "")), self._me.mri))]
        )
        chats = [self._chat(c) for c in convs]
        return sorted(chats, key=lambda c: c.last_time or datetime.min, reverse=True)

    def _chat(self, conv: dict[str, Any]) -> Chat:
        conv_id = str(conv["id"])
        kind = rules.chat_kind(conv) or "group"
        last = conv.get("lastMessage") or {}
        preview = ""
        if last and not rules.is_system(last):
            preview = one_line(rules.message_text(last)[0], 90)
        return Chat(
            ref=self._refs.put(BackendId("c", conv_id)),
            id=conv_id,
            kind=kind,
            title=one_line(self._title(conv, kind), 80),
            last_time=rules.parse_time(
                last.get("originalarrivaltime") or conv.get("properties", {}).get("lastimreceivedtime")
            ),
            last_from=one_line(str(last.get("imdisplayname") or ""), 40) if preview else "",
            last_preview=preview,
            unread=rules.is_unread(conv, self._me),
        )

    def _title(self, conv: dict[str, Any], kind: ChatKind) -> str:
        """Chat title; 1:1 and untitled group chats are named after the other people (cached per chat)."""
        props = conv.get("threadProperties") or {}
        conv_id = str(conv["id"])
        if kind == "channel":
            team = str(props.get("spaceThreadTopic") or "")
            topic = str(props.get("topicThreadTopic") or props.get("topic") or "")
            if team and topic and team != topic:
                return f"{team} / {topic}"
            return team or topic or "General"
        title = str(props.get("topicThreadTopic") or props.get("topic") or "")
        if title:
            return title
        cached = self._profiles.get(f"chat:{conv_id}")
        if cached and cached.name:
            return cached.name
        others = [m for m in self._members(conv_id) if not self._is_me(m)]
        names = [m.name for m in others if m.name]
        if kind == "chat":
            title = names[0] if names else ""
        elif names:
            title = ", ".join(n.split()[0] for n in names[:4]) + (
                f" +{len(others) - 4}" if len(others) > 4 else ""
            )
        if not title:
            return "1:1 chat" if kind == "chat" else "Group chat"
        self._profiles.put_all([Member(name=title, mri=f"chat:{conv_id}")])
        return title

    def _is_me(self, member: Member) -> bool:
        """You, by user ID or by name (old chats can hold your user ID from before a tenant move)."""
        me = self._me
        return member.mri == me.mri or (bool(member.name) and member.name.casefold() == me.name.casefold())

    def _members(self, conv_id: str) -> list[Member]:
        """Chat members with names: from the directory, else from the names on their messages."""
        thread = self._session.request(CHAT, f"/v1/threads/{_q(conv_id)}?view=msnp24Equivalent") or {}
        mris = [
            str(m.get("id")) for m in thread.get("members") or [] if str(m.get("id", "")).startswith("8:")
        ]
        self._load_profiles(mris)
        members = [self._profiles.get(mri) or Member(name="", mri=mri) for mri in mris]
        if any(not m.name for m in members):
            seen = self._names_on_messages(conv_id)
            self._profiles.put_all(
                [Member(name=seen[m.mri], mri=m.mri) for m in members if not m.name and m.mri in seen]
            )
            members = [self._profiles.get(m.mri) or m for m in members]
        return members

    def _names_on_messages(self, conv_id: str) -> dict[str, str]:
        path = (
            f"/v1/users/ME/conversations/{_q(conv_id)}/messages?view={_MESSAGE_VIEW}&pageSize=50&startTime=1"
        )
        reply = self._session.request(CHAT, path) or {}
        names: dict[str, str] = {}
        for m in reply.get("messages") or []:
            name = str(m.get("imdisplayname") or m.get("fromDisplayNameInToken") or "")
            if name and not rules.is_system(m):
                names.setdefault(rules.mri_of(str(m.get("from", ""))), name)
        return names

    def _load_profiles(self, mris: list[str]) -> None:
        """Look up unknown users; users the directory cannot resolve (other tenants) get no name."""
        missing = sorted({m for m in mris if m and self._profiles.get(m) is None})
        for start in range(0, len(missing), 50):
            batch = missing[start : start + 50]
            reply = self._session.request(
                MIDDLE_TIER, _PROFILE_PATH.format("false"), method="POST", body=batch, auth=(SPACES,)
            )
            found = _profiles_of(reply)
            known = {m.mri for m in found}
            self._profiles.put_all(found + [Member(name="", mri=m) for m in batch if m not in known])

    def _profile_by_email(self, email: str) -> Member:
        reply = self._session.request(
            MIDDLE_TIER, _PROFILE_PATH.format("true"), method="POST", body=[email], auth=(SPACES,)
        )
        found = _profiles_of(reply)
        if not found:
            raise InvalidInputError(f"No Teams user with email {email}.")
        self._profiles.put_all(found)
        return found[0]

    def _resolve(self, chat: str) -> Chat:
        text = chat.strip()
        if text.startswith("c:"):
            conv_id = self._refs.get(text, kinds="c").item_id
        elif text.startswith(("19:", "48:")):
            conv_id = text
        else:
            return rules.pick_one(self._all_chats(), text, "chat", lambda c: c.title)
        known = next((c for c in self._conversations if c.get("id") == conv_id), None)
        if known is None:
            known = self._session.request(
                CHAT, f"/v1/users/ME/conversations/{_q(conv_id)}?view=msnp24Equivalent"
            )
            if not isinstance(known, dict) or "id" not in known:
                raise NotFoundError(f"Teams chat {chat} not found (deleted, or you left it).")
        return self._chat(known)

    def read(
        self, chat: str, *, limit: int = 20, since: datetime | None = None, older: str | None = None
    ) -> MessagePage:
        """See `TeamsApi.read`."""
        target = self._resolve(chat)
        if older:
            cursor = self._refs.get(older, kinds="o")
            if cursor.container_id != target.id:
                raise InvalidInputError("This cursor belongs to another chat.")
            path = cursor.item_id
        else:
            path = (
                f"/v1/users/ME/conversations/{_q(target.id)}/messages?view={_MESSAGE_VIEW}"
                f"&pageSize={min(max(limit, 1), 200)}&startTime=1"
            )
        reply = self._session.request(CHAT, path) or {}
        raw = [m for m in reply.get("messages") or [] if not rules.is_system(m)]
        messages = [self._message(m, target.id) for m in raw]
        if since:
            messages = [m for m in messages if m.time is None or m.time >= since]
        reached_since = since is not None and len(messages) < len(raw)
        older_ref = None
        link = self._own_path(str((reply.get("_metadata") or {}).get("backwardLink") or ""))
        if link and not reached_since:
            older_ref = self._refs.put(BackendId("o", link, target.id))
        return MessagePage(chat=target, messages=list(reversed(messages)), older=older_ref)

    def _own_path(self, link: str) -> str:
        """The path of a paging link, only when it points to our own chat service.

        Paging links are full URLs from the server; we only follow them on the chat-service host so a
        request with your tokens never goes anywhere else.
        """
        if not link:
            return ""
        base = str((self._session.status() or {}).get("chatService") or "")
        if not base or not link.startswith(base + "/v1/"):
            return ""
        return link[len(base) :]

    def _message(self, raw: dict[str, Any], conv_id: str) -> TeamsMessage:
        props = raw.get("properties") or {}
        text, mentions, quote, images = rules.message_text(raw)
        sender_mri = rules.mri_of(str(raw.get("from", "")))
        sender = str(raw.get("imdisplayname") or raw.get("fromDisplayNameInToken") or "")
        if not sender and (profile := self._profiles.get(sender_mri)):
            sender = profile.name
        thread = str(raw.get("conversationLink") or "").rsplit("/conversations/", 1)[-1] or conv_id
        return TeamsMessage(
            ref=self._refs.put(BackendId("p", str(raw.get("id")), thread)),
            id=str(raw.get("id")),
            sender=one_line(sender or "Unknown", 60),
            sender_mri=sender_mri,
            time=rules.parse_time(raw.get("originalarrivaltime") or raw.get("composetime")),
            text="" if props.get("deletetime") else text,
            edited=bool(props.get("edittime") or raw.get("edittime")),
            deleted=bool(props.get("deletetime")),
            quote=quote,
            mentions=mentions,
            reactions=rules.reactions_of(props),
            files=rules.files_of(props),
            images=images,
        )

    def draft(
        self,
        body: str,
        *,
        chat: str | None = None,
        to: str | None = None,
        reply_to: str | None = None,
        body_format: BodyFormat = "markdown",
    ) -> TeamsDraft:
        """See `TeamsApi.draft`."""
        rules.check_body(body)
        body = with_signoff(body, self._ctx.settings.signoff)
        quote = ""
        target: Chat | None = None
        person: Member | None = None
        reply_id = reply_mri = reply_name = reply_quote = ""
        if reply_to:
            ref = self._refs.get(reply_to, kinds="p")
            conv_id = ref.container_id.split(";", 1)[0]
            target = self._resolve(conv_id)
            original = (
                self._session.request(
                    CHAT, f"/v1/users/ME/conversations/{_q(conv_id)}/messages/{ref.item_id}"
                )
                or {}
            )
            message = self._message(original, conv_id)
            reply_id, reply_mri, reply_name = message.id, message.sender_mri, message.sender
            reply_quote = f"{message.sender}: {one_line(message.text, 120)}"
            if target.kind != "channel":
                quote = rules.quote_html(message.id, message.sender_mri, message.sender, message.text)
            if chat and self._resolve(chat).id != target.id:
                raise InvalidInputError("reply_to belongs to another chat than `chat`.")
        elif to:
            target, person = self._chat_with(to)
        elif chat:
            target = self._resolve(chat)
        else:
            raise InvalidInputError("Give chat, to, or reply_to.")
        members = self._draft_members(target, person) if "@[" in body else []
        mentions = rules.resolve_mentions(body, members)
        content, text = rules.body_html(body, body_format, mentions)
        draft = TeamsDraft(
            id=uuid.uuid4().hex,
            chat_id=target.id if target else "",
            chat_title=target.title if target else (person.name if person else ""),
            chat_kind=target.kind if target else "chat",
            to=person,
            html=quote + content,
            text=text,
            reply_to_id=reply_id,
            reply_quote=reply_quote,
            reply_to_mri=reply_mri,
            reply_to_name=reply_name,
            mentions=mentions,
            created=now(),
        )
        self._drafts.put(draft)
        return draft.model_copy(update={"ref": self._refs.put(BackendId("d", draft.id))})

    def _chat_with(self, to: str) -> tuple[Chat | None, Member | None]:
        """The 1:1 chat with a person (name or email); for an email without a chat, the person."""
        one_on_ones = [c for c in self._all_chats() if c.kind == "chat"]
        if "@" not in to:
            try:
                return rules.pick_one(one_on_ones, to, "1:1 chat", lambda c: c.title), None
            except InvalidInputError as exc:
                raise InvalidInputError(f"{exc} To start a new chat, give the person's email.") from exc
        person = self._profile_by_email(to.strip())
        for conv in self._conversations:
            if rules.other_member_mri(str(conv.get("id", "")), self._me.mri) == person.mri:
                return self._chat(conv), None
        return None, person

    def _draft_members(self, target: Chat | None, person: Member | None) -> list[Member]:
        if target is None:
            return [m for m in (person, self._me) if m]
        return self._members(target.id)

    def send(self, draft_ref: str, *, send_at: datetime | None = None) -> str:
        """See `TeamsApi.send`."""
        draft = self._drafts.get(self._refs.get(draft_ref, kinds="d").item_id)
        settings = self._ctx.settings
        rules.check_protected(draft.chat_title, settings.teams_protected)
        if send_at is not None and send_at <= now():
            raise InvalidInputError("send_at is in the past.")
        self._ctx.audit.guard("teams_send")
        chat_id = draft.chat_id or self._start_chat(draft)
        target = chat_id
        if draft.chat_kind == "channel" and draft.reply_to_id:
            target = f"{chat_id};messageid={draft.reply_to_id}"
        client_id = rules.client_message_id()
        payload = rules.message_payload(
            draft.html, display_name=self._me.name, mentions=draft.mentions, client_id=client_id
        )
        if send_at is None:
            reply = (
                self._session.request(
                    CHAT, f"/v1/users/ME/conversations/{_q(target)}/messages", method="POST", body=payload
                )
                or {}
            )
        else:
            reply = (
                self._session.request(
                    CHAT,
                    "/v1/users/ME/drafts",
                    method="POST",
                    body=rules.scheduled_payload(target, payload, send_at),
                )
                or {}
            )
        message_id = str(reply.get("OriginalArrivalTime") or "")
        when = f" (scheduled for {send_at:%Y-%m-%d %H:%M})" if send_at else ""
        self._ctx.audit.record("teams_send", f"{draft.chat_title}{when}: {one_line(draft.text, 80)}")
        self._drafts.remove(draft.id)
        ref = self._refs.put(BackendId("p", message_id, target)) if message_id else ""
        verb = "Scheduled" if send_at else "Sent"
        return f"{verb} to {draft.chat_title} ({draft.chat_kind}){when}." + (
            f" Message {ref}." if ref else ""
        )

    def _start_chat(self, draft: TeamsDraft) -> str:
        if draft.to is None:
            raise InvalidInputError("The draft has no chat.")
        body = {"members": [{"mri": draft.to.mri, "role": "Admin"}]}
        reply = self._session.request(CHAT, "/v1/users/ME/conversations", method="PUT", body=body) or {}
        conv_id = str(reply.get("id") or "")
        if not conv_id:
            raise BackendUnavailableError("Teams did not create the chat; nothing was sent.")
        return conv_id

    def status(self, query: str) -> Presence:
        """See `TeamsApi.status`."""
        card = self._resolve_contact(query)
        reply = self._session.request(
            PRESENCE,
            "/v1/presence/getpresence/",
            method="POST",
            body=[{"mri": card["mri"]}],
            auth=(SPACES,),
        )
        rows = reply if isinstance(reply, list) else []
        raw = next((r.get("presence") or {} for r in rows if r.get("mri") == card["mri"]), {})
        return Presence(
            name=card["name"],
            mri=card["mri"],
            email=card["email"],
            title=card["title"],
            department=card["department"],
            office=card["office"],
            **rules.parse_presence(raw),
        )

    def mark_read(self, chat: str) -> str:
        """See `TeamsApi.mark_read`."""
        target = self._resolve(chat)
        last = self._last_message_id(target.id)
        if not last:
            return f"{target.title}: no messages to mark."
        self._session.request(
            CHAT,
            f"/v1/users/ME/conversations/{_q(target.id)}/properties?name=consumptionhorizon",
            method="PUT",
            body={"consumptionhorizon": rules.consumption_horizon(last)},
        )
        return f"Marked {target.title} read."

    def _last_message_id(self, conv_id: str) -> str:
        path = (
            f"/v1/users/ME/conversations/{_q(conv_id)}/messages?view={_MESSAGE_VIEW}&pageSize=1&startTime=1"
        )
        reply = self._session.request(CHAT, path) or {}
        messages = reply.get("messages") or []
        return str(messages[0].get("id")) if messages else ""

    def digest(self, *, unread_only: bool = True, limit: int = 15) -> Digest:
        """See `TeamsApi.digest`."""
        unread = [c for c in self._all_chats() if c.unread][:limit]
        path = (
            f"/v1/users/ME/conversations/{_MENTIONS}/messages?view={_MESSAGE_VIEW}"
            f"&pageSize={max(limit * 2, 20)}&startTime=1"
        )
        try:
            reply = self._session.request(CHAT, path) or {}
        except M365Error:
            reply = {}
        mentions: list[Mention] = []
        for raw in reply.get("messages") or []:
            item = rules.parse_activity(raw)
            if item is None or (unread_only and not item["unread"]):
                continue
            ref = self._refs.put(BackendId("p", item["message_id"], item["thread"]))
            mentions.append(
                Mention(
                    ref=ref,
                    who=one_line(item["who"], 50),
                    where=one_line(item["where"], 50),
                    kind=item["kind"],
                    preview=item["preview"],
                    time=item["time"],
                    unread=item["unread"],
                )
            )
            if len(mentions) >= limit:
                break
        return Digest(mentions=mentions, unread=unread)

    def search(self, query: str, *, offset: int = 0, limit: int = 15) -> Page[SearchHit]:
        """See `TeamsApi.search`."""
        if not query.strip():
            raise InvalidInputError("Give something to search for.")
        body = {
            "Scenario": {"Name": "powerbar"},
            "Cvid": str(uuid.uuid4()),
            "EntityRequests": [
                {
                    "EntityType": "Message",
                    "ContentSources": ["Teams"],
                    "Provenances": ["Teams"],
                    "Query": {"QueryString": query, "DisplayQueryString": query},
                    "From": offset,
                    "Size": min(max(limit, 1), 50) + 1,
                }
            ],
        }
        try:
            reply = self._session.request(
                SUBSTRATE,
                "/search/api/v1/suggestions?scenario=powerbar",
                method="POST",
                body=body,
                auth=(SUBSTRATE_OFFICE,),
            )
        except M365Error:
            return Page(items=[], offset=offset, has_more=False)
        hits: list[SearchHit] = []
        for group in (reply or {}).get("Groups") or []:
            for s in group.get("Suggestions") or []:
                thread = str(s.get("ClientThreadId") or "")
                if not thread.startswith("19:") and not thread.startswith("48:"):
                    continue
                team, channel = str(s.get("Team") or ""), str(s.get("Channel") or "")
                where = f"{team} / {channel}" if team and channel else team or channel
                hits.append(
                    SearchHit(
                        chat_ref=self._refs.put(BackendId("c", thread)),
                        kind=rules.search_kind(str(s.get("ThreadType") or "")),
                        who=one_line(rules.sender_name(s.get("From")), 50),
                        where=one_line(where, 50),
                        preview=one_line(str(s.get("Preview") or ""), 160),
                        time=rules.parse_time(s.get("DateTimeSent")),
                    )
                )
        return Page(items=hits[:limit], offset=offset, has_more=len(hits) > limit)

    def files(self, chat: str, *, limit: int = 30) -> list[SharedFile]:
        """See `TeamsApi.files`."""
        found: list[SharedFile] = []
        for raw, entry in self._file_messages(chat):
            found.append(
                SharedFile(
                    name=entry["name"],
                    file_type=entry["type"],
                    who=one_line(str(raw.get("imdisplayname") or ""), 40),
                    time=rules.parse_time(raw.get("originalarrivaltime") or raw.get("composetime")),
                )
            )
            if len(found) >= limit:
                break
        return found

    def file_text(self, chat: str, name: str) -> FileText:
        """See `TeamsApi.file_text`."""
        entries = [entry for _, entry in self._file_messages(chat)]
        match = rules.pick_one(entries, name, "shared file", lambda e: e["name"])
        data = self._download_file(match)
        import tempfile

        suffix = "." + match["name"].rsplit(".", 1)[-1] if "." in match["name"] else ""
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
            handle.write(data)
            temp = Path(handle.name)
        try:
            text, cut = truncate(clean_body(extract_text(temp)), self._ctx.settings.body_max_chars)
        finally:
            temp.unlink(missing_ok=True)
        return FileText(name=match["name"], text=text, truncated=cut)

    def _download_file(self, entry: dict[str, str]) -> bytes:
        """Download a shared file's bytes from SharePoint, using the file's own drive."""
        base = entry.get("base", "").rstrip("/")
        url = entry["url"]
        if not base or not url.startswith(base):
            raise InvalidInputError(f"Cannot locate {entry['name']} in SharePoint.")
        relative = urllib.parse.unquote(url[len(base) :].lstrip("/"))
        for prefix in ("Documents/", "Shared Documents/"):
            relative = relative.removeprefix(prefix)
        path = f"/_api/v2.0/drive/root:/{urllib.parse.quote(relative)}:/content"
        return self._session.download(base, path, auth=(SHAREPOINT_AUD,))

    def _file_messages(self, chat: str) -> Iterator[tuple[dict[str, Any], dict[str, str]]]:
        """Yield (message, file-entry) pairs for files in a chat's recent messages, newest first."""
        target = self._resolve(chat)
        path = (
            f"/v1/users/ME/conversations/{_q(target.id)}/messages?view={_MESSAGE_VIEW}"
            f"&pageSize={_FILE_SCAN}&startTime=1"
        )
        reply = self._session.request(CHAT, path) or {}
        for raw in reply.get("messages") or []:
            for entry in rules.file_entries(raw.get("properties") or {}):
                yield raw, entry

    def update_message(
        self,
        message_ref: str,
        *,
        react: str | None = None,
        remove_react: bool = False,
        edit: str | None = None,
        delete: bool = False,
        body_format: BodyFormat = "markdown",
    ) -> str:
        """See `TeamsApi.update_message`."""
        ref = self._refs.get(message_ref, kinds="p")
        conv_id = ref.container_id.split(";", 1)[0]
        base = f"/v1/users/ME/conversations/{_q(conv_id)}/messages/{ref.item_id}"
        chosen = [bool(react or remove_react), edit is not None, delete]
        if sum(chosen) != 1:
            raise InvalidInputError("Give exactly one of: react, edit, delete.")
        if delete:
            self._ctx.audit.guard("teams_delete")
            self._session.request(CHAT, f"{base}?behavior=softDelete", method="DELETE")
            self._ctx.audit.record("teams_delete", f"message {message_ref}")
            return "Message deleted."
        if edit is not None:
            rules.check_body(edit)
            self._ctx.audit.guard("teams_edit")
            content, _ = rules.body_html(edit, body_format, [])
            self._session.request(
                CHAT, base, method="PUT", body=rules.edit_payload(content, ref.item_id, self._me.name)
            )
            self._ctx.audit.record("teams_edit", f"message {message_ref}")
            return "Message edited."
        key = rules.reaction_key(react or "")
        self._ctx.audit.guard("teams_react")
        if remove_react:
            self._session.request(
                CHAT,
                f"{base}/properties?name=emotions",
                method="DELETE",
                body={"emotions": {"key": key, "value": ref.item_id}},
            )
            self._ctx.audit.record("teams_react", f"removed {key} on {message_ref}")
            return f"Removed reaction {key}."
        self._session.request(
            CHAT,
            f"{base}/properties?name=emotions",
            method="PUT",
            body={"emotions": {"key": key, "value": ref.item_id}},
        )
        self._ctx.audit.record("teams_react", f"{key} on {message_ref}")
        return f"Reacted {key}."

    def _resolve_contact(self, query: str) -> dict[str, str]:
        """Resolve a colleague by name or email to an MRI and directory card via the people directory."""
        if "@" in query:
            card = self._card_by_email(query.strip())
            if card:
                return card
        people = self._search_people(query)
        if not people:
            raise InvalidInputError(f"No Teams user matches {query!r}. Try a full name or an email.")
        exact = [p for p in people if query.casefold() in (p["name"].casefold(), p["email"].casefold())]
        chosen = exact or people
        if len(chosen) > 1:
            names = ", ".join(dict.fromkeys(p["name"] for p in chosen))
            raise InvalidInputError(f"{query!r} matches several people: {names}. Use a full name or email.")
        return chosen[0]

    def _search_people(self, query: str) -> list[dict[str, str]]:
        body = {
            "Scenario": {"Name": "peoplepicker.newChat"},
            "Cvid": str(uuid.uuid4()),
            "EntityRequests": [
                {
                    "Query": {"QueryString": query, "DisplayQueryString": query},
                    "EntityType": "People",
                    "Provenances": ["Mailbox", "Directory"],
                    "From": 0,
                    "Size": 10,
                }
            ],
        }
        try:
            reply = self._session.request(
                SUBSTRATE,
                "/search/api/v1/suggestions?scenario=peoplepicker.newChat",
                method="POST",
                body=body,
                auth=(SUBSTRATE_AUD,),
            )
        except M365Error:
            return []
        out: list[dict[str, str]] = []
        for group in (reply or {}).get("Groups") or []:
            for s in group.get("Suggestions") or []:
                mri = str(s.get("MRI") or "")
                if not mri.startswith("8:"):
                    continue
                emails = s.get("EmailAddresses") or []
                out.append(
                    {
                        "mri": mri,
                        "name": str(s.get("DisplayName") or s.get("Text") or ""),
                        "email": str(emails[0] if emails else s.get("QueryText") or ""),
                        "title": str(s.get("JobTitle") or ""),
                        "department": str(s.get("Department") or ""),
                        "office": str(s.get("OfficeLocation") or s.get("AdditionalOfficeLocation") or ""),
                    }
                )
        return out

    def _card_by_email(self, email: str) -> dict[str, str] | None:
        reply = self._session.request(
            MIDDLE_TIER, _PROFILE_PATH.format("true"), method="POST", body=[email], auth=(SPACES,)
        )
        values = reply.get("value") if isinstance(reply, dict) else None
        card = (values or [None])[0]
        if not isinstance(card, dict) or not card.get("mri"):
            return None
        return {
            "mri": str(card["mri"]),
            "name": str(card.get("displayName") or card.get("givenName") or ""),
            "email": str(card.get("email") or card.get("userPrincipalName") or email),
            "title": str(card.get("jobTitle") or ""),
            "department": str(card.get("department") or ""),
            "office": str(card.get("userLocation") or ""),
        }


def _profiles_of(reply: Any) -> list[Member]:
    values = reply.get("value") if isinstance(reply, dict) else reply
    out = []
    for p in values if isinstance(values, list) else []:
        if isinstance(p, dict) and p.get("mri"):
            out.append(
                Member(
                    name=str(p.get("displayName") or p.get("givenName") or ""),
                    mri=str(p["mri"]),
                    email=str(p.get("email") or p.get("userPrincipalName") or ""),
                )
            )
    return out
