from __future__ import annotations

import base64
import json
import re
import time
import urllib.parse
from datetime import UTC, datetime
from typing import Any, ClassVar

from m365_cli.backends.browser.session import PageSession, retry_wait
from m365_cli.core.errors import BackendUnavailableError, InvalidInputError, M365Error, NotFoundError

OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/"
OUTLOOK_HOSTS = ("outlook.cloud.microsoft", "outlook.office.com", "outlook.office365.com")
OUTLOOK_AUD = "https://outlook.office.com"
REST = "/api/beta"
_MAX_RETRIES = 3
_FOLDER_TTL = 600
_UPLOAD_CHUNK = 10 * 320 * 1024
_WALL_TIME = re.compile(r"^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})")


class OutlookAuthError(BackendUnavailableError):
    """Outlook on the web refused the session (expired or missing sign-in, or no access)."""


class OutlookApiError(M365Error):
    """Outlook on the web answered with an unexpected status or an error message."""

    def __init__(self, message: str, status: int = 0) -> None:
        """Store the HTTP status next to the message."""
        super().__init__(message)
        self.status = status


def rest_id(ews_id: str) -> str:
    """REST form of an Exchange item id (URL-safe base64)."""
    return ews_id.replace("/", "-").replace("+", "_")


def ews_id(rest: str) -> str:
    """Exchange (service.svc) form of a REST item id."""
    return rest.replace("-", "/").replace("_", "+")


def from_utc(value: str | None) -> datetime | None:
    """A UTC timestamp of the API ('2026-11-17T08:22:45Z') as naive local time."""
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return (parsed if parsed.tzinfo else parsed.replace(tzinfo=UTC)).astimezone().replace(tzinfo=None)


def to_utc(value: datetime) -> str:
    """Naive local time as an API UTC timestamp."""
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def from_wall(value: dict[str, Any] | None) -> datetime | None:
    """An API date-time-with-zone object, read in the mailbox time zone, as a naive wall-clock time."""
    m = _WALL_TIME.match(str((value or {}).get("DateTime", "")))
    return datetime.fromisoformat(m.group(1)) if m else None


def wall(value: datetime, zone: str) -> dict[str, str]:
    """Naive wall-clock time in the mailbox time zone as an API date-time-with-zone object."""
    return {"DateTime": value.strftime("%Y-%m-%dT%H:%M:%S"), "TimeZone": zone}


def quote(value: str) -> str:
    """Quote a value for a single-quoted OData string literal."""
    return value.replace("'", "''")


class OutlookWebSession(PageSession):
    """Calls the Outlook REST API (/api/beta) and Outlook on the web's service.svc from inside its tab.

    Item ids are immutable ids (they survive moves and sending). Event times are read and written in
    the mailbox time zone; message timestamps come as UTC and are turned into local time.
    """

    app: ClassVar[str] = "Outlook"
    url: ClassVar[str] = OUTLOOK_URL
    hosts: ClassVar[tuple[str, ...]] = OUTLOOK_HOSTS
    auth_error: ClassVar[type[BackendUnavailableError]] = OutlookAuthError

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        """Remember the browser settings; nothing starts before the first call."""
        super().__init__(*args, **kwargs)
        self._cache: dict[str, Any] = {}

    def _is_ready(self, state: dict[str, Any]) -> bool:
        return bool(state.get("me")) and OUTLOOK_AUD in (state.get("bearers") or {})

    def request(
        self,
        path: str,
        *,
        method: str = "GET",
        body: Any = None,
        headers: dict[str, str] | None = None,
        binary: bool = False,
        raw: bytes | None = None,
        token: bool = True,
        service: str = "self",
    ) -> Any:
        """Call a path on the Outlook origin with the Outlook token; parsed JSON (or bytes with `binary`).

        `raw` sends bytes as the body. Without `token` no Authorization header is sent (for URLs that
        carry their own token). `service` is "self" (the tab's origin) or an https:// origin.

        429 answers are retried with the server's Retry-After; one 401 reloads Outlook to renew the token.

        Raises:
            OutlookAuthError: Unauthorized or forbidden.
            NotFoundError: 404.
            OutlookApiError: Any other unexpected status.
        """
        payload: dict[str, Any] = {
            "service": service,
            "path": path,
            "method": method,
            "body": body,
            "auth": [OUTLOOK_AUD] if token else [],
            "headers": headers or {},
            "binary": binary,
        }
        if raw is not None:
            payload["bodyBase64"] = base64.b64encode(raw).decode()
        renewed = False
        for attempt in range(_MAX_RETRIES + 1):
            reply = self._call(payload)
            status = int(reply.get("status") or 0)
            if status == 0:
                error = str(reply.get("error", "no answer"))
                if error.startswith("no-token") and not renewed:
                    self._renew()
                    renewed = True
                    continue
                raise BackendUnavailableError(f"Outlook request failed: {error}")
            if status == 401 and not renewed:
                self._renew()
                renewed = True
                continue
            if status == 429 and attempt < _MAX_RETRIES:
                time.sleep(retry_wait(reply.get("retryAfter"), attempt))
                continue
            if binary and 200 <= status < 300:
                return base64.b64decode(str(reply.get("base64") or ""))
            return _result(status, str(reply.get("body") or ""), method, path)
        raise OutlookApiError("Outlook is rate-limiting requests (429). Wait a minute, then retry.", 429)

    def rest(
        self,
        path: str,
        *,
        method: str = "GET",
        body: Any = None,
        mailbox: str | None = None,
        timezone: str | None = None,
        text: bool = False,
        binary: bool = False,
    ) -> Any:
        """Outlook REST call; `path` follows /me (or /users/{mailbox} for another mailbox).

        Times in the answer use `timezone` (default UTC); bodies are HTML unless `text`.
        """
        who = f"/users/{urllib.parse.quote(mailbox)}" if mailbox else "/me"
        prefer = [
            'IdType="ImmutableId"',
            f'outlook.timezone="{timezone or "UTC"}"',
            f'outlook.body-content-type="{"text" if text else "html"}"',
        ]
        headers = {"Prefer": ", ".join(prefer)}
        if mailbox:
            headers["X-AnchorMailbox"] = f"SMTP:{mailbox}"
        if body is not None:
            headers["Content-Type"] = "application/json"
        return self.request(f"{REST}{who}{path}", method=method, body=body, headers=headers, binary=binary)

    def upload(self, upload_url: str, data: bytes) -> None:
        """Upload bytes to an upload session in chunks (the URL carries its own token)."""
        parts = urllib.parse.urlsplit(upload_url)
        service = "self" if parts.hostname in OUTLOOK_HOSTS else f"https://{parts.hostname}"
        for start in range(0, len(data), _UPLOAD_CHUNK):
            chunk = data[start : start + _UPLOAD_CHUNK]
            self.request(
                f"{parts.path}?{parts.query}",
                method="PUT",
                raw=chunk,
                headers={"Content-Range": f"bytes {start}-{start + len(chunk) - 1}/{len(data)}"},
                token=False,
                service=service,
            )

    def rest_pages(
        self, path: str, *, max_items: int, mailbox: str | None = None, **kwargs: Any
    ) -> list[dict]:
        """Items of a listing, following the server's next links until `max_items`."""
        items: list[dict] = []
        page = self.rest(path, mailbox=mailbox, **kwargs) or {}
        while True:
            items += page.get("value", [])
            link = page.get("@odata.nextLink")
            if not link or len(items) >= max_items:
                return items[:max_items]
            parts = urllib.parse.urlsplit(link)
            headers = {
                "Prefer": f'IdType="ImmutableId", outlook.timezone="{kwargs.get("timezone") or "UTC"}"'
            }
            if mailbox:
                headers["X-AnchorMailbox"] = f"SMTP:{mailbox}"
            page = self.request(f"{parts.path}?{parts.query}", headers=headers) or {}

    def svc(self, action: str, body: dict[str, Any], *, mailbox: str | None = None) -> dict[str, Any]:
        """Run one Outlook on the web service.svc action; returns its first response message.

        Raises:
            OutlookApiError: The action failed.
        """
        request = {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": self.timezone(),
                    },
                },
            },
            "Body": {"__type": f"{action}Request:#Exchange", **body},
        }
        reply = self.svc_raw(action, request, mailbox=mailbox) or {}
        answer = reply.get("Body") or {}
        items = (answer.get("ResponseMessages") or {}).get("Items") or []
        if not items:
            detail = answer.get("FaultMessage") or answer.get("ExceptionName") or "no answer"
            raise OutlookApiError(f"Outlook could not run {action}: {detail}")
        first = items[0]
        if first.get("ResponseClass") == "Error":
            raise OutlookApiError(
                f"Outlook could not run {action}: {first.get('MessageText') or first.get('ResponseCode')}"
            )
        return first

    def svc_raw(self, action: str, request: Any, *, mailbox: str | None = None) -> Any:
        """Send a raw service.svc request body; returns the parsed answer."""
        headers = {
            "Action": action,
            "Content-Type": "application/json; charset=utf-8",
            "Prefer": 'IdType="ImmutableId"',
            "x-req-source": "Mail",
        }
        if mailbox:
            headers["X-AnchorMailbox"] = f"SMTP:{mailbox}"
        return self.request(
            f"/owa/service.svc?action={action}&app=Mail", method="POST", body=request, headers=headers
        )

    def cloud_settings(self, names: str, *, account: bool = False) -> list[dict[str, Any]]:
        """Outlook cloud settings by name (signature names; with `account`, one signature's content)."""
        path = "/ows/v1/OutlookCloudSettings/settings/" + ("account" if account else "")
        headers = {"x-islargesetting": "true"} if account else {}
        reply = self.request(f"{path}?settingname={urllib.parse.quote(names)}", headers=headers)
        return [s for s in reply if isinstance(s, dict)] if isinstance(reply, list) else []

    def user_options(self) -> dict[str, Any]:
        """Outlook on the web user options (compose font and more), cached."""
        if "options" not in self._cache:
            reply = self.svc_raw("GetOwaUserConfiguration", {}) or {}
            self._cache["options"] = reply.get("UserOptions") or {}
        return self._cache["options"]

    def timezone(self) -> str:
        """The mailbox time zone (Windows name, e.g. 'Central Europe Standard Time'), cached."""
        if "tz" not in self._cache:
            reply = self.rest("/mailboxsettings?$select=TimeZone") or {}
            self._cache["tz"] = reply.get("TimeZone") or "UTC"
        return self._cache["tz"]

    def me(self) -> tuple[str, str]:
        """(display name, primary address) of the signed-in user, cached."""
        if "me" not in self._cache:
            reply = self.rest("") or {}
            self._cache["me"] = (
                str(reply.get("DisplayName") or ""),
                str(reply.get("EmailAddress") or "").lower(),
            )
        return self._cache["me"]

    def folders(self, mailbox: str | None = None) -> list[dict[str, Any]]:
        """Mail folders of a mailbox with their path ('Inbox/Alerts'), cached for a few minutes.

        Each entry has Id, DisplayName, WellKnownName (lower case, may be empty) and path.
        """
        key = f"folders|{mailbox or ''}"
        cached = self._cache.get(key)
        if cached and time.monotonic() - cached[0] < _FOLDER_TTL:
            return cached[1]
        found: list[dict[str, Any]] = []

        def walk(path: str, prefix: str) -> None:
            select = "$top=250&$select=Id,DisplayName,ChildFolderCount,WellKnownName"
            for f in self.rest_pages(f"{path}?{select}", max_items=2000, mailbox=mailbox):
                label = f"{prefix}/{f.get('DisplayName')}" if prefix else str(f.get("DisplayName") or "")
                found.append({**f, "WellKnownName": str(f.get("WellKnownName") or "").lower(), "path": label})
                if f.get("ChildFolderCount"):
                    walk(f"/mailfolders/{f['Id']}/childfolders", label)

        walk("/mailfolders", "")
        self._cache[key] = (time.monotonic(), found)
        return found

    def shared_mailboxes(self) -> list[tuple[str, str]]:
        """(display name, address) of the other mailboxes added to Outlook on the web, cached."""
        if "shared" not in self._cache:
            reply = self.svc_raw("GetOtherMailboxConfiguration", {}) or {}
            self._cache["shared"] = [
                (
                    str(e.get("DisplayName") or e.get("PrincipalSMTPAddress")),
                    str(e["PrincipalSMTPAddress"]).lower(),
                )
                for e in reply.get("OtherMailboxEntries") or []
                if e.get("PrincipalSMTPAddress")
            ]
        return self._cache["shared"]

    def mailbox(self, name: str | None) -> str | None:
        """Address of another mailbox given by address or (part of its) name; None for your own.

        Raises:
            NotFoundError: No unique mailbox with that name.
        """
        key = (name or "").strip().lower()
        if not key or key == self.me()[1]:
            return None
        if "@" in key:
            return key
        boxes = self.shared_mailboxes()
        found = [a for n, a in boxes if n.lower() == key] or [a for n, a in boxes if key in n.lower()]
        if len(found) == 1:
            return found[0]
        known = ", ".join(f"{n} <{a}>" for n, a in boxes) or "none"
        raise NotFoundError(f"No unique mailbox {name!r}. Mailboxes: {known}.")

    def find_people(self, query: str, *, top: int = 10) -> list[dict[str, Any]]:
        """People matching a name or address (your contacts and the directory), best first."""
        search = urllib.parse.quote(f'"{query.replace(chr(34), "")}"')
        reply = self.rest(f"/people?$search={search}&$top={top}") or {}
        return list(reply.get("value", []))

    def resolve(self, value: str) -> tuple[str, str] | None:
        """(name, address) for an address or a name that matches one person; None when not unique."""
        text = value.strip()
        m = re.match(r"^(.*?)\s*<([^<>@\s]+@[^<>\s]+)>$", text)
        if m:
            return m.group(1).strip() or m.group(2), m.group(2)
        if "@" in text:
            return text, text
        people = [p for p in self.find_people(text) if _addresses(p)]
        exact = [p for p in people if str(p.get("DisplayName", "")).lower() == text.lower()]
        chosen = exact if exact else people
        if len(chosen) != 1:
            return None
        return str(chosen[0].get("DisplayName") or ""), _addresses(chosen[0])[0]

    def schedules(
        self, addresses: list[str], start: datetime, end: datetime, minutes: int
    ) -> list[tuple[str, str | None]]:
        """Free/busy of people as (address, availability digits or None when not available).

        One digit per `minutes` from `start` (wall-clock time in the mailbox time zone): 0 free,
        1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
        """
        zone = self.timezone()
        body = {
            "Schedules": addresses,
            "StartTime": wall(start, zone),
            "EndTime": wall(end, zone),
            "AvailabilityViewInterval": minutes,
        }
        reply = self.rest("/calendar/getschedule", method="POST", body=body, timezone=zone) or {}
        result = []
        for schedule in reply.get("value", []):
            view = str(schedule.get("AvailabilityView") or "")
            result.append(
                (str(schedule.get("ScheduleId") or ""), None if schedule.get("Error") or not view else view)
            )
        return result

    def immutable_ids(self, ids: list[str], *, mailbox: str | None = None) -> list[str]:
        """Immutable ids for ids that a search returned (search answers with ordinary ids)."""
        if not ids:
            return []
        body = {"InputIds": ids, "SourceIdType": "RestId", "TargetIdType": "RestImmutableEntryId"}
        reply = self.rest("/translateexchangeids", method="POST", body=body, mailbox=mailbox) or {}
        out = [str(r.get("TargetId") or "") for r in reply.get("value", [])]
        return [new or old for new, old in zip(out, ids, strict=False)] if len(out) == len(ids) else ids


def _addresses(person: dict[str, Any]) -> list[str]:
    return [str(e["Address"]) for e in person.get("EmailAddresses") or [] if e.get("Address")]


def _result(status: int, text: str, method: str, path: str) -> Any:
    if 200 <= status < 300:
        if not text.strip():
            return None
        try:
            return json.loads(text)
        except ValueError:
            return text
    where = f"{method} {path.split('?', 1)[0]}"
    detail = _error_text(text)
    if status in (401, 403):
        raise OutlookAuthError(f"Outlook refused {where} ({status}): {detail or 'no access'}")
    if status == 404:
        raise NotFoundError(f"Outlook item not found ({detail or where}); list again to get current refs.")
    if status == 400:
        raise InvalidInputError(f"Outlook rejected {where}: {detail}")
    raise OutlookApiError(f"Outlook answered {status} for {where}: {detail}", status)


def _error_text(text: str) -> str:
    try:
        data = json.loads(text)
    except ValueError:
        return text.strip().replace("\n", " ")[:200]
    if isinstance(data, dict):
        error = data.get("error") or {}
        body = data.get("Body") or {}
        return str(error.get("message") or body.get("FaultMessage") or body.get("ExceptionName") or "")[:300]
    return ""
