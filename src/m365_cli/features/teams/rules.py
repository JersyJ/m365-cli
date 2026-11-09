from __future__ import annotations

import fnmatch
import html
import json
import re
import secrets
from collections.abc import Callable, Iterable, Sequence
from datetime import UTC, datetime
from html.parser import HTMLParser
from typing import Any

from m365_cli.core.errors import InvalidInputError
from m365_cli.core.markup import BodyFormat, to_html, to_text
from m365_cli.core.text import one_line, sanitize
from m365_cli.features.teams.models import ChatKind, Member

KIND_BY_PRODUCT: dict[str, ChatKind] = {
    "OneToOneChat": "chat",
    "Chat": "group",
    "Meeting": "meeting",
    "TeamsStandardChannel": "channel",
    "TeamsPrivateChannel": "channel",
    "TeamsSharedChannel": "channel",
    "TeamsTeam": "channel",
}
KIND_BY_THREAD: dict[str, ChatKind] = {
    "chat": "group",
    "meeting": "meeting",
    "topic": "channel",
    "space": "channel",
}
NO_EXPIRY_MS = 253402300800000
MESSAGE_TYPES = ("RichText/Html", "Text", "RichText")
REPLY_SCHEMA = "http://schema.skype.com/Reply"
MENTION_SCHEMA = "http://schema.skype.com/Mention"
_MENTION_TOKEN = re.compile(r"@\[([^\]\n]{1,80})\]")
_PLACEHOLDER = "M365MENTION{}X"
_ONE_ON_ONE = re.compile(r"^19:([0-9a-f-]{36})_([0-9a-f-]{36})@unq\.gbl\.spaces$", re.I)
_BLOCKS = {"p", "div", "br", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "pre", "table"}
_SYSTEM_PREFIXES = ("ThreadActivity/", "Event/", "RichText/Media_")


def parse_time(value: object) -> datetime | None:
    """Teams time (ISO 8601 in UTC, or epoch milliseconds) as a naive local datetime."""
    if value in (None, ""):
        return None
    try:
        if isinstance(value, int | float) or str(value).isdigit():
            moment = datetime.fromtimestamp(int(str(value).split(".", 1)[0]) / 1000, tz=UTC)
        else:
            moment = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except (ValueError, OverflowError, OSError):
        return None
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=UTC)
    return moment.astimezone().replace(tzinfo=None)


def mri_of(value: str) -> str:
    """The user ID ('8:orgid:<guid>') from a 'from' URL like 'https://…/contacts/8:orgid:<guid>'."""
    return value.rsplit("/", 1)[-1] if value else ""


def chat_kind(conversation: dict[str, Any]) -> ChatKind | None:
    """Kind of a conversation from the chat service, or None for system streams and unknown types."""
    if str(conversation.get("id", "")).startswith("48:"):
        return None
    props = conversation.get("threadProperties") or {}
    product = str(props.get("productThreadType") or "")
    if product:
        return KIND_BY_PRODUCT.get(product)
    if _ONE_ON_ONE.match(str(conversation.get("id", ""))):
        return "chat"
    return KIND_BY_THREAD.get(str(props.get("threadType") or ""))


def is_listed(conversation: dict[str, Any], *, now_ms: int) -> bool:
    """False for system streams, deleted or hidden threads, and channels you have left.

    Teams keeps listing deleted teams (`isdeleted`) and left channels (expired membership); reading them
    returns 404 "soft-deleted", so they are hidden. Meeting chats with an expired membership stay
    readable, so only channels are checked for expiry.
    """
    if chat_kind(conversation) is None:
        return False
    props = conversation.get("threadProperties") or {}
    if _truthy(props.get("isdeleted")) or _truthy(props.get("hidden")):
        return False
    if chat_kind(conversation) != "channel":
        return True
    member = conversation.get("memberProperties") or {}
    expires = int(member.get("memberExpirationTime") or 0)
    return not (0 < expires < NO_EXPIRY_MS and expires < now_ms)


def _truthy(value: object) -> bool:
    return value is True or str(value).lower() == "true"


def other_member_mri(conversation_id: str, me_mri: str) -> str:
    """The other person's user ID in a 1:1 chat ('' when it cannot be told from the ID).

    The ID of a 1:1 chat holds both user GUIDs: 19:<guid>_<guid>@unq.gbl.spaces. After a tenant
    move, old chats hold your previous ID instead of the current one; then this returns ''.
    """
    match = _ONE_ON_ONE.match(conversation_id)
    if not match:
        return ""
    me = me_mri.rsplit(":", 1)[-1].lower()
    first, second = match.group(1).lower(), match.group(2).lower()
    if me not in (first, second):
        return ""
    return f"8:orgid:{second if first == me else first}"


def is_mine(message: dict[str, Any], me: Member) -> bool:
    """True when you wrote the message: by user ID, or by display name for chats from before a tenant move."""
    if mri_of(str(message.get("from", ""))) == me.mri:
        return True
    name = str(message.get("imdisplayname") or message.get("fromDisplayNameInToken") or "")
    return bool(name) and name.casefold() == me.name.casefold()


def is_unread(conversation: dict[str, Any], me: Member) -> bool:
    """True when the newest message is from someone else and newer than your read position.

    The read position is the first field of `consumptionhorizon` ('<message id>;<time>;<client id>').
    """
    last = conversation.get("lastMessage") or {}
    if not last or is_mine(last, me):
        return False
    if str(last.get("messagetype", "")).startswith(_SYSTEM_PREFIXES):
        return False
    horizon = str((conversation.get("properties") or {}).get("consumptionhorizon") or "")
    read_id = horizon.split(";", 1)[0]
    try:
        return int(str(last.get("id") or 0)) > int(read_id or 0)
    except ValueError:
        return False


def is_system(message: dict[str, Any]) -> bool:
    """True for events (member added, call ended, recording) instead of written messages."""
    kind = str(message.get("messagetype", ""))
    return kind.startswith(_SYSTEM_PREFIXES) or (kind not in MESSAGE_TYPES and kind != "")


def reactions_of(properties: dict[str, Any]) -> dict[str, int]:
    """Reaction counts from `properties.emotions`.

    Teams sends `emotions` either as a JSON string or as a list. The 'follow' entry is a thread
    subscription, not a reaction; entries without users are reactions that were removed.
    """
    raw = properties.get("emotions")
    if isinstance(raw, str):
        try:
            raw = json.loads(raw) if raw.strip() else []
        except ValueError:
            raw = []
    counts: dict[str, int] = {}
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict) or item.get("key") in (None, "", "follow"):
            continue
        users = item.get("users") or []
        if users:
            counts[str(item["key"])] = len(users)
    return counts


def files_of(properties: dict[str, Any]) -> list[str]:
    """Names of the files attached to a message (`properties.files`, a JSON string)."""
    raw = properties.get("files")
    if isinstance(raw, str):
        try:
            raw = json.loads(raw) if raw.strip() else []
        except ValueError:
            raw = []
    names = []
    for item in raw if isinstance(raw, list) else []:
        if isinstance(item, dict):
            names.append(str(item.get("fileName") or item.get("title") or "file"))
    return names


class _TextParser(HTMLParser):
    """Turns Teams message HTML into plain text, and collects mentions, the quoted reply and images."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.mentions: list[str] = []
        self.images = 0
        self.quote_sender = ""
        self.quote_preview: list[str] = []
        self._quote_depth = 0
        self._in_sender = False
        self._mention: list[str] | None = None
        self._href: str | None = None
        self._link_text: list[str] = []
        self._list_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        a = dict(attrs)
        if tag == "blockquote" and (a.get("itemtype") == REPLY_SCHEMA or self._quote_depth):
            self._quote_depth += 1
            return
        if self._quote_depth:
            self._in_sender = tag == "strong" and a.get("itemprop") == "mri"
            return
        if a.get("itemtype") == MENTION_SCHEMA and tag == "span":
            self._mention = []
        elif tag == "img":
            alt = a.get("alt") or ""
            if a.get("itemtype") == "http://schema.skype.com/Emoji" or "emoji" in (a.get("class") or ""):
                self.parts.append(alt)
            else:
                self.images += 1
                self.parts.append("[image]")
        elif tag == "emoji":
            self.parts.append(a.get("alt") or "")
        elif tag == "a":
            self._href = a.get("href") or ""
            self._link_text = []
        elif tag in ("ul", "ol"):
            self._list_depth += 1
        elif tag == "li":
            self.parts.append("\n" + "  " * max(self._list_depth - 1, 0) + "- ")
        elif tag in _BLOCKS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if self._quote_depth:
            if tag == "blockquote":
                self._quote_depth -= 1
            elif tag == "strong":
                self._in_sender = False
            elif tag in _BLOCKS:
                self.quote_preview.append(" ")
            return
        if tag == "span" and self._mention is not None:
            name = "".join(self._mention).strip()
            self.mentions.append(name)
            self.parts.append(f"@{name}")
            self._mention = None
        elif tag == "a" and self._href is not None:
            text = "".join(self._link_text).strip()
            href = self._href
            self.parts.append(text if not href or href in (text, f"mailto:{text}") else f"{text} ({href})")
            self._href = None
        elif tag in ("ul", "ol"):
            self._list_depth = max(self._list_depth - 1, 0)
            self.parts.append("\n")
        elif tag in _BLOCKS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self._quote_depth:
            if self._in_sender:
                self.quote_sender += data
            else:
                self.quote_preview.append(data)
        elif self._mention is not None:
            self._mention.append(data)
        elif self._href is not None:
            self._link_text.append(data)
        else:
            self.parts.append(data)


def html_to_text(content: str) -> tuple[str, list[str], str, int]:
    """Plain text of a Teams HTML message.

    Returns (text, mentioned names, quote as 'Sender: preview', number of images).
    """
    parser = _TextParser()
    parser.feed(content or "")
    parser.close()
    text = re.sub(r"[ \t\xa0]+", " ", "".join(parser.parts))
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    quote = ""
    preview = one_line("".join(parser.quote_preview), 160)
    if parser.quote_sender or preview:
        quote = f"{one_line(parser.quote_sender, 60)}: {preview}" if parser.quote_sender else preview
    return sanitize(text), parser.mentions, quote, parser.images


def message_text(message: dict[str, Any]) -> tuple[str, list[str], str, int]:
    """Text, mentions, quote and image count of a raw message (plain 'Text' messages have no HTML)."""
    content = str(message.get("content") or "")
    if str(message.get("messagetype")) == "Text" and "<" not in content:
        return sanitize(html.unescape(content)).strip(), [], "", 0
    return html_to_text(content)


def matches(text: str, query: str) -> bool:
    """Case-insensitive match of every word of `query` in `text`."""
    hay = text.casefold()
    return all(word in hay for word in query.casefold().split())


def pick_one[T](items: Sequence[T], query: str, label: str, name: Callable[[T], str]) -> T:
    """The single item whose name matches `query`; an exact match wins over partial matches.

    Raises:
        InvalidInputError: No match or more than one match (the error lists the candidates).
    """
    exact = [i for i in items if name(i).casefold() == query.casefold()]
    found = exact or [i for i in items if matches(name(i), query)]
    if len(found) == 1:
        return found[0]
    if not found:
        raise InvalidInputError(f"No {label} matches {query!r}.")
    names = ", ".join(sorted({name(i) for i in found})[:6])
    raise InvalidInputError(f"{query!r} matches several {label}s: {names}. Be more specific or use a ref.")


def resolve_mentions(body: str, members: Iterable[Member]) -> list[Member]:
    """Members named by `@[Name]` in the body, in order; each name must match one chat member.

    Raises:
        InvalidInputError: A name matches no member or several members.
    """
    pool = list(members)
    found: list[Member] = []
    for match in _MENTION_TOKEN.finditer(body):
        member = pick_one(pool, match.group(1).strip(), "chat member", lambda m: m.name)
        found.append(member)
    return found


def body_html(body: str, body_format: BodyFormat, mentions: Sequence[Member]) -> tuple[str, str]:
    """Teams HTML and plain text of the body; `@[Name]` becomes a real mention of that member.

    Markdown is converted without raw HTML (`<tags>` in the text stay text).
    """
    index = 0

    def placeholder(_: re.Match[str]) -> str:
        nonlocal index
        token = _PLACEHOLDER.format(index)
        index += 1
        return token

    marked = _MENTION_TOKEN.sub(placeholder, body) if mentions else body
    rendered = to_html(marked, body_format)
    plain = to_text(body, body_format)
    for i, member in enumerate(mentions):
        rendered = rendered.replace(_PLACEHOLDER.format(i), mention_html(member.name, i), 1)
        plain = plain.replace(f"@[{member.name}]", f"@{member.name}")
    for match in _MENTION_TOKEN.finditer(plain):
        plain = plain.replace(match.group(0), f"@{match.group(1)}")
    return rendered, plain


def mention_html(name: str, index: int) -> str:
    """HTML of one person mention, as the Teams web client writes it."""
    span = f'<span itemtype="{MENTION_SCHEMA}" itemscope itemid="{index}">{html.escape(name)}</span>'
    return (
        f'<readonly class="skipProofing" itemtype="{MENTION_SCHEMA}" contenteditable="false" '
        f'spellcheck="false">{span}</readonly>'
    )


def mentions_property(mentions: Sequence[Member]) -> str:
    """The `properties.mentions` JSON string that goes with `mention_html`."""
    return json.dumps(
        [
            {
                "@type": MENTION_SCHEMA,
                "itemid": str(i),
                "mri": m.mri,
                "mentionType": "person",
                "displayName": m.name,
            }
            for i, m in enumerate(mentions)
        ],
        ensure_ascii=False,
    )


def quote_html(message_id: str, sender_mri: str, sender_name: str, preview: str) -> str:
    """The quoted-reply block Teams shows above a reply."""
    short = one_line(preview, 200)
    return (
        f'<blockquote itemscope="" itemtype="{REPLY_SCHEMA}" itemid="{html.escape(message_id)}">'
        f'<strong itemprop="mri" itemid="{html.escape(sender_mri)}">{html.escape(sender_name)}</strong>'
        f'<span itemprop="time" itemid="{html.escape(message_id)}"></span>'
        f'<p itemprop="preview">{html.escape(short)}</p></blockquote>'
    )


def client_message_id() -> str:
    """A new random client message ID (19 digits, like the Teams client)."""
    return str(secrets.randbelow(9 * 10**18) + 10**18)


def message_payload(
    content_html: str, *, display_name: str, mentions: Sequence[Member], client_id: str
) -> dict[str, Any]:
    """Body of a chat-service send.

    `imdisplayname` is required: without it the sender shows as blank. The server answers with
    `OriginalArrivalTime`, which is the real message ID (the echoed `id` is only `client_id`).
    """
    return {
        "content": content_html,
        "messagetype": "RichText/Html",
        "contenttype": "text",
        "imdisplayname": display_name,
        "clientmessageid": client_id,
        "properties": {"importance": "", "subject": "", "mentions": mentions_property(mentions)},
    }


def scheduled_payload(conversation_id: str, message: dict[str, Any], send_at: datetime) -> dict[str, Any]:
    """Body of a scheduled send: a 'ScheduledDraft' that the Teams server sends at `send_at`."""
    moment = send_at.astimezone(UTC)
    stamp = moment.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    created = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    inner = dict(message)
    inner["properties"] = dict(message["properties"]) | {
        "cards": "[]",
        "links": "[]",
        "files": "[]",
        "formatVariant": "TEAMS",
        "draftId": message["clientmessageid"],
    }
    inner |= {
        "id": "-1",
        "type": "Message",
        "conversationid": conversation_id,
        "composetime": created,
        "originalarrivaltime": created,
        "draftDetails": {"sendAt": stamp},
        "threadtype": "streamofdrafts",
        "innerThreadId": conversation_id,
    }
    return {
        "draftDetails": {"sendAt": str(int(moment.timestamp() * 1000))},
        "draftType": "ScheduledDraft",
        "innerThreadId": conversation_id,
        "message": inner,
    }


def check_protected(title: str, patterns: Sequence[str]) -> None:
    """Refuse to send to a chat whose title matches a `teams_protected` pattern ('*' wildcards).

    Raises:
        InvalidInputError: The chat is protected.
    """
    for pattern in patterns:
        if fnmatch.fnmatchcase(title.casefold(), pattern.casefold()):
            raise InvalidInputError(
                f"Chat {title!r} is protected (teams_protected = {pattern!r}); nothing was sent."
            )


def check_body(body: str) -> None:
    """Refuse an empty message or one above the Teams size limit.

    Raises:
        InvalidInputError: Empty or too long.
    """
    if not body.strip():
        raise InvalidInputError("The message is empty.")
    if len(body) > 25_000:
        raise InvalidInputError("The message is too long for Teams (limit about 28 KB); shorten it.")
