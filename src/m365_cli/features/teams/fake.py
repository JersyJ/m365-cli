from __future__ import annotations

import uuid
from datetime import datetime
from typing import TYPE_CHECKING

from m365_cli.backends.fake.state import FakeChat, FakeState, FakeTeamsMessage, Outgoing
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.core.markup import BodyFormat
from m365_cli.core.paging import Page, take_page
from m365_cli.core.refs import BackendId
from m365_cli.core.text import one_line, with_signoff
from m365_cli.core.timeutil import now
from m365_cli.features.teams import rules
from m365_cli.features.teams.models import (
    Chat,
    ChatKind,
    Digest,
    Member,
    Mention,
    MessagePage,
    Presence,
    SearchHit,
    TeamsDraft,
    TeamsMessage,
)

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

_PAGE_CURSOR = "older-than:"


def _mri(name: str) -> str:
    return "8:orgid:" + name.casefold().replace(" ", "-")


class FakeTeamsService:
    """Teams operations on the fake backend; sends only go to `FakeState.outbox`."""

    def __init__(self, ctx: AppContext) -> None:
        """Use the context's fake store, refs, send log and settings."""
        self._ctx = ctx
        self._store = ctx.fake
        self._refs = ctx.refs

    @property
    def _state(self) -> FakeState:
        return self._store.state

    def _me(self) -> Member:
        me = self._state.me
        return Member(name=me.name, mri=_mri(me.name), email=me.email)

    def _member(self, name: str) -> Member:
        person = next((p for p in self._state.people if p.name == name), None)
        email = person.email if person else self._state.me.email if name == self._state.me.name else ""
        return Member(name=name, mri=_mri(name), email=email)

    def _chat(self, c: FakeChat) -> Chat:
        visible = [m for m in c.messages if not m.system]
        last = visible[-1] if visible else None
        return Chat(
            ref=self._refs.put(BackendId("c", c.id)),
            id=c.id,
            kind=c.kind,
            title=c.title,
            last_time=last.time if last else None,
            last_from=last.sender if last else "",
            last_preview=one_line(rules.html_to_text(last.html)[0], 90) if last else "",
            unread=c.unread,
        )

    def _all(self) -> list[Chat]:
        chats = [self._chat(c) for c in self._state.teams_chats if not c.deleted]
        return sorted(chats, key=lambda c: c.last_time or datetime.min, reverse=True)

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
        found = self._all()
        if query:
            found = [c for c in found if rules.matches(c.title, query)]
        if kind:
            found = [c for c in found if c.kind == kind]
        if unread_only:
            found = [c for c in found if c.unread]
        items, more = take_page(found, offset=offset, limit=limit)
        return Page(items=items, offset=offset, has_more=more)

    def _fake_chat(self, chat_id: str) -> FakeChat:
        found = next((c for c in self._state.teams_chats if c.id == chat_id and not c.deleted), None)
        if found is None:
            raise NotFoundError(f"Teams chat {chat_id} not found (deleted, or you left it).")
        return found

    def _resolve(self, chat: str) -> FakeChat:
        text = chat.strip()
        if text.startswith("c:"):
            return self._fake_chat(self._refs.get(text, kinds="c").item_id)
        if text.startswith(("19:", "48:")):
            return self._fake_chat(text)
        picked = rules.pick_one(self._all(), text, "chat", lambda c: c.title)
        return self._fake_chat(picked.id)

    def _message(self, m: FakeTeamsMessage, chat_id: str) -> TeamsMessage:
        text, mentions, quote, images = rules.html_to_text(m.html)
        return TeamsMessage(
            ref=self._refs.put(BackendId("p", m.id, chat_id)),
            id=m.id,
            sender=m.sender,
            sender_mri=_mri(m.sender),
            time=m.time,
            text="" if m.deleted else text,
            edited=m.edited,
            deleted=m.deleted,
            quote=quote,
            mentions=mentions,
            reactions=dict(m.reactions),
            files=list(m.files),
            images=images,
        )

    def read(
        self, chat: str, *, limit: int = 20, since: datetime | None = None, older: str | None = None
    ) -> MessagePage:
        """See `TeamsApi.read`."""
        c = self._resolve(chat)
        visible = [m for m in c.messages if not m.system]
        if older:
            cursor = self._refs.get(older, kinds="o")
            if cursor.container_id != c.id:
                raise InvalidInputError("This cursor belongs to another chat.")
            before = cursor.item_id.removeprefix(_PAGE_CURSOR)
            visible = visible[: next((i for i, m in enumerate(visible) if m.id == before), len(visible))]
        page = visible[-limit:]
        if since:
            page = [m for m in page if m.time >= since]
        older_ref = None
        if len(visible) > limit and page and (since is None or len(page) == limit):
            older_ref = self._refs.put(BackendId("o", _PAGE_CURSOR + page[0].id, c.id))
        return MessagePage(
            chat=self._chat(c), messages=[self._message(m, c.id) for m in page], older=older_ref
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
        target: FakeChat | None = None
        person: Member | None = None
        quote = reply_quote = ""
        reply: TeamsMessage | None = None
        if reply_to:
            ref = self._refs.get(reply_to, kinds="p")
            target = self._fake_chat(ref.container_id.split(";", 1)[0])
            original = next((m for m in target.messages if m.id == ref.item_id), None)
            if original is None:
                raise NotFoundError(f"Message {reply_to} not found.")
            reply = self._message(original, target.id)
            reply_quote = f"{reply.sender}: {one_line(reply.text, 120)}"
            if target.kind != "channel":
                quote = rules.quote_html(reply.id, reply.sender_mri, reply.sender, reply.text)
            if chat and self._resolve(chat).id != target.id:
                raise InvalidInputError("reply_to belongs to another chat than `chat`.")
        elif to:
            target, person = self._chat_with(to)
        elif chat:
            target = self._resolve(chat)
        else:
            raise InvalidInputError("Give chat, to, or reply_to.")
        if target is not None:
            members = [self._member(n) for n in target.members]
        else:
            members = [m for m in (person, self._me()) if m]
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
            reply_to_id=reply.id if reply else "",
            reply_quote=reply_quote,
            reply_to_mri=reply.sender_mri if reply else "",
            reply_to_name=reply.sender if reply else "",
            mentions=mentions,
            created=now(),
        )
        with self._store.edit() as state:
            state.teams_drafts[draft.id] = draft.model_dump(mode="json")
        return draft.model_copy(update={"ref": self._refs.put(BackendId("d", draft.id))})

    def _chat_with(self, to: str) -> tuple[FakeChat | None, Member | None]:
        one_on_ones = [c for c in self._state.teams_chats if c.kind == "chat" and not c.deleted]
        if "@" not in to:
            try:
                return rules.pick_one(one_on_ones, to, "1:1 chat", lambda c: c.title), None
            except InvalidInputError as exc:
                raise InvalidInputError(f"{exc} To start a new chat, give the person's email.") from exc
        person = next((p for p in self._state.people if p.email.casefold() == to.strip().casefold()), None)
        if person is None:
            raise InvalidInputError(f"No Teams user with email {to}.")
        existing = next((c for c in one_on_ones if person.name in c.members), None)
        return (existing, None) if existing else (None, self._member(person.name))

    def send(self, draft_ref: str, *, send_at: datetime | None = None) -> str:
        """See `TeamsApi.send`."""
        draft_id = self._refs.get(draft_ref, kinds="d").item_id
        raw = self._state.teams_drafts.get(draft_id)
        if raw is None:
            raise NotFoundError("Unknown Teams draft (already sent, or older than 7 days). Draft it again.")
        draft = TeamsDraft.model_validate(raw)
        rules.check_protected(draft.chat_title, self._ctx.settings.teams_protected)
        if send_at is not None and send_at <= now():
            raise InvalidInputError("send_at is in the past.")
        self._ctx.audit.guard("teams_send")
        with self._store.edit() as state:
            chat_id = draft.chat_id
            if not chat_id and draft.to is not None:
                chat_id = state.new_id("19:new-")
                state.teams_chats.append(
                    FakeChat(
                        id=chat_id, kind="chat", title=draft.to.name, members=[draft.to.name, state.me.name]
                    )
                )
            message_id = state.new_id("")
            if send_at is None:
                chat = next(c for c in state.teams_chats if c.id == chat_id)
                chat.messages.append(
                    FakeTeamsMessage(id=message_id, sender=state.me.name, time=now(), html=draft.html)
                )
            state.outbox.append(
                Outgoing(
                    at=now(),
                    action="teams_send",
                    item_id=chat_id,
                    summary=draft.text,
                    recipients=[draft.chat_title],
                    send_at=send_at,
                )
            )
            state.teams_drafts.pop(draft_id, None)
        when = f" (scheduled for {send_at:%Y-%m-%d %H:%M})" if send_at else ""
        self._ctx.audit.record("teams_send", f"{draft.chat_title}{when}: {one_line(draft.text, 80)}")
        ref = self._refs.put(BackendId("p", message_id, chat_id))
        verb = "Scheduled" if send_at else "Sent"
        return f"{verb} to {draft.chat_title} ({draft.chat_kind}){when}. Message {ref}."

    def status(self, query: str) -> Presence:
        """See `TeamsApi.status`."""
        q = query.strip().casefold()
        people = [*self._state.people, self._state.me]
        exact = [p for p in people if q in (p.name.casefold(), p.email.casefold())]
        found = exact or [p for p in people if rules.matches(p.name, query)]
        if not found:
            raise InvalidInputError(f"No Teams user matches {query!r}. Try a full name or an email.")
        if len(found) > 1:
            names = ", ".join(dict.fromkeys(p.name for p in found))
            raise InvalidInputError(f"{query!r} matches several people: {names}. Use a full name or email.")
        person = found[0]
        status = person.teams_status or "Available"
        return Presence(
            name=person.name,
            mri=_mri(person.name),
            email=person.email,
            title=person.title,
            department=person.department,
            office=person.office,
            status=status,
            available=status == "Available",
            out_of_office="office" in status.casefold() or bool(person.teams_ooo),
            ooo_note=person.teams_ooo,
            message=person.teams_note,
        )

    def mark_read(self, chat: str) -> str:
        """See `TeamsApi.mark_read`."""
        c = self._resolve(chat)
        with self._store.edit() as state:
            fc = next(x for x in state.teams_chats if x.id == c.id)
            fc.unread = False
        return f"Marked {c.title} read."

    def digest(self, *, unread_only: bool = True, limit: int = 15) -> Digest:
        """See `TeamsApi.digest`."""
        me_first = self._state.me.name.split()[0].casefold()
        unread = [self._chat(c) for c in self._state.teams_chats if c.unread and not c.deleted][:limit]
        mentions: list[Mention] = []
        for c in self._state.teams_chats:
            if c.deleted:
                continue
            for m in c.messages:
                names = [n.casefold() for n in rules.html_to_text(m.html)[1]]
                if me_first not in names:
                    continue
                mentions.append(
                    Mention(
                        ref=self._refs.put(BackendId("p", m.id, c.id)),
                        who=m.sender,
                        where=c.title,
                        kind="mentioned you",
                        preview=one_line(rules.html_to_text(m.html)[0], 160),
                        time=m.time,
                        unread=c.unread,
                    )
                )
        if unread_only:
            mentions = [x for x in mentions if x.unread]
        mentions = sorted(mentions, key=lambda x: x.time or datetime.min, reverse=True)[:limit]
        return Digest(mentions=mentions, unread=unread)

    def search(self, query: str, *, offset: int = 0, limit: int = 15) -> Page[SearchHit]:
        """See `TeamsApi.search`."""
        if not query.strip():
            raise InvalidInputError("Give something to search for.")
        hits: list[SearchHit] = []
        for c in self._state.teams_chats:
            if c.deleted:
                continue
            for m in c.messages:
                text = rules.html_to_text(m.html)[0]
                if not m.deleted and rules.matches(text, query):
                    hits.append(
                        SearchHit(
                            chat_ref=self._refs.put(BackendId("c", c.id)),
                            kind=c.kind,
                            who=m.sender,
                            where=c.title if c.kind == "channel" else "",
                            preview=one_line(text, 160),
                            time=m.time,
                        )
                    )
        hits.sort(key=lambda h: h.time or datetime.min, reverse=True)
        items, more = take_page(hits, offset=offset, limit=limit)
        return Page(items=items, offset=offset, has_more=more)

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
        chosen = [bool(react or remove_react), edit is not None, delete]
        if sum(chosen) != 1:
            raise InvalidInputError("Give exactly one of: react, edit, delete.")
        chat = self._fake_chat(ref.container_id.split(";", 1)[0])
        with self._store.edit() as state:
            fc = next(x for x in state.teams_chats if x.id == chat.id)
            msg = next((m for m in fc.messages if m.id == ref.item_id), None)
            if msg is None:
                raise NotFoundError(f"Message {message_ref} not found.")
            if delete:
                self._ctx.audit.guard("teams_delete")
                msg.deleted = True
                msg.html = ""
                action, summary, result = "teams_delete", f"message {message_ref}", "Message deleted."
            elif edit is not None:
                rules.check_body(edit)
                self._ctx.audit.guard("teams_edit")
                content, _ = rules.body_html(edit, body_format, [])
                msg.html = content
                msg.edited = True
                action, summary, result = "teams_edit", f"message {message_ref}", "Message edited."
            else:
                key = rules.reaction_key(react or "")
                self._ctx.audit.guard("teams_react")
                if remove_react:
                    msg.reactions.pop(key, None)
                    action, summary, result = "teams_react", f"removed {key}", f"Removed reaction {key}."
                else:
                    msg.reactions[key] = msg.reactions.get(key, 0) + 1
                    action, summary, result = "teams_react", f"{key} on {message_ref}", f"Reacted {key}."
            state.outbox.append(
                Outgoing(at=now(), action=action, item_id=chat.id, summary=summary, recipients=[chat.title])
            )
        self._ctx.audit.record(action, summary)
        return result
