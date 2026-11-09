from __future__ import annotations

import concurrent.futures as cf
import contextlib
import threading
from collections.abc import Callable, Iterator
from datetime import datetime
from typing import Any

import pythoncom
import pywintypes
import win32com.client

from m365_cli.backends.outlook.constants import (
    EXCHANGE_PUBLIC_FOLDERS,
    EXCHANGE_USER_TYPES,
    MAPI_E_NOT_FOUND,
    OlFolder,
    OlItemType,
)
from m365_cli.core.errors import (
    BackendUnavailableError,
    InvalidInputError,
    M365Error,
    NotFoundError,
    UnresolvedRecipientsError,
)
from m365_cli.core.refs import BackendId
from m365_cli.core.timeutil import to_dasl_utc

type Com = Any
"""Untyped COM dispatch object (pywin32 has no type information for COM)."""

FOLDER_ALIASES: dict[str, OlFolder] = {
    "inbox": OlFolder.INBOX,
    "sent": OlFolder.SENT,
    "drafts": OlFolder.DRAFTS,
    "deleted": OlFolder.DELETED,
    "trash": OlFolder.DELETED,
    "junk": OlFolder.JUNK,
    "outbox": OlFolder.OUTBOX,
    "calendar": OlFolder.CALENDAR,
}

SYSTEM_FOLDER_NAMES = frozenset(
    {
        "conversation action settings",
        "conversation history",
        "rss feeds",
        "sync issues",
        "yammer root",
        "files",
        "quick step settings",
        "social activity notifications",
        "externalcontacts",
        "persondata",
    }
)
"""Folders of type mail that hold no real mail; searching everywhere skips them."""


def iterate(items: Com) -> Iterator[Com]:
    """Iterate an Outlook Items collection (GetFirst/GetNext is much faster than indexing)."""
    item = items.GetFirst()
    while item is not None:
        yield item
        item = items.GetNext()


def com_error_info(exc: pywintypes.com_error) -> tuple[int, str]:
    """(hresult, message) from a com_error, whose args are (hresult, text, excepinfo, argerror).

    When present, excepinfo is (wcode, source, description, helpfile, helpcontext, scode); its
    description and scode are more precise than text and hresult.
    """
    args = (*exc.args, None, None, None)
    hresult, text, excepinfo = args[0], args[1], args[2]
    if excepinfo:
        return int(excepinfo[5] or hresult or 0), str(excepinfo[2] or text or "").strip()
    return int(hresult or 0), str(text or "").strip()


class Outlook:
    """Thin helper over the Outlook Application/Namespace. Only use it on the session's COM thread."""

    def __init__(self) -> None:
        """Connect to classic Outlook.

        Raises:
            BackendUnavailableError: No signed-in classic Outlook profile.
        """
        self.app: Com = win32com.client.Dispatch("Outlook.Application")
        self.ns: Com = self.app.GetNamespace("MAPI")
        try:
            self.default_store_id: str = self.ns.GetDefaultFolder(OlFolder.INBOX).StoreID
        except pywintypes.com_error as exc:
            raise BackendUnavailableError(
                f"Classic Outlook is not connected ({com_error_info(exc)[1]}). "
                "Open classic Outlook and sign in."
            ) from exc

    def mailboxes(self) -> list[tuple[str, str]]:
        """Mailboxes in the Outlook profile as (name, kind); kind is primary | account | shared | other."""
        kinds = {0: "primary", 4: "account", 1: "shared"}
        return [
            (str(store.DisplayName), kinds.get(store.ExchangeStoreType, "other"))
            for store in self.ns.Stores
            if store.ExchangeStoreType != EXCHANGE_PUBLIC_FOLDERS
        ]

    def _store(self, mailbox: str) -> Com | None:
        """Store of a mailbox in the profile (exact or partial name); None when not in the profile."""
        key = mailbox.strip().lower()
        stores = [s for s in self.ns.Stores if s.ExchangeStoreType != EXCHANGE_PUBLIC_FOLDERS]
        exact = [s for s in stores if str(s.DisplayName).lower() == key]
        partial = exact or [s for s in stores if key in str(s.DisplayName).lower()]
        return partial[0] if len(partial) == 1 else None

    def default_folder(self, folder: OlFolder, mailbox: str | None = None) -> Com:
        """Default folder of the default mailbox, or of `mailbox` (in the profile or shared with you).

        Raises:
            NotFoundError: The mailbox is unknown or you have no access to it.
        """
        if not mailbox:
            return self.ns.GetDefaultFolder(folder)
        store = self._store(mailbox)
        try:
            if store is not None:
                with contextlib.suppress(pywintypes.com_error):
                    return store.GetDefaultFolder(folder)
            return self.ns.GetSharedDefaultFolder(self.recipient(mailbox), folder)
        except (pywintypes.com_error, UnresolvedRecipientsError) as exc:
            known = ", ".join(name for name, _ in self.mailboxes())
            reason = com_error_info(exc)[1] if isinstance(exc, pywintypes.com_error) else str(exc)
            raise NotFoundError(f"No access to mailbox {mailbox!r} ({reason}). Known: {known}.") from exc

    def folder(self, path: str, mailbox: str | None = None) -> Com:
        """'inbox', 'sent', … or a path below the mailbox root such as 'Inbox/Projects/X'.

        Raises:
            NotFoundError: A part of the path does not exist.
        """
        key = path.strip().strip("/")
        if alias := FOLDER_ALIASES.get(key.lower()):
            return self.default_folder(alias, mailbox)
        current: Com = self.default_folder(OlFolder.INBOX, mailbox).Parent
        for part in key.split("/"):
            match = next((f for f in current.Folders if f.Name.lower() == part.lower()), None)
            if match is None:
                raise NotFoundError(f"Folder not found: {path!r} (missing '{part}').")
            current = match
        return current

    def calendar_between(self, start: datetime, end: datetime, owner: str | None = None) -> Com:
        """Calendar items overlapping [start, end), recurring occurrences expanded, sorted by start.

        `owner` is a colleague or mailbox whose calendar is shared with you.

        Raises:
            InvalidInputError: The owner's calendar is not shared with you.
        """
        try:
            items = self.default_folder(OlFolder.CALENDAR, owner).Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")
            found = items.Restrict(
                f"@SQL=\"urn:schemas:calendar:dtstart\" < '{to_dasl_utc(end)}' "
                f"AND \"urn:schemas:calendar:dtend\" > '{to_dasl_utc(start)}'"
            )
            if owner:
                found.GetFirst()
            return found
        except (NotFoundError, pywintypes.com_error) as exc:
            if not owner:
                raise
            raise InvalidInputError(
                f"{owner}'s calendar is not shared with you; use cal_find_free_time or people_lookup for "
                "their free/busy instead."
            ) from exc

    def mail_folders(self, mailbox: str | None = None) -> list[tuple[str, Com]]:
        """All mail folders of a mailbox as ('Inbox/Alerts', folder), for searching everywhere.

        Deleted Items, Junk, Outbox, Drafts and system folders (sync issues, RSS, …) are left out.
        """
        skip_ids = set()
        for f in (OlFolder.DELETED, OlFolder.JUNK, OlFolder.OUTBOX, OlFolder.DRAFTS):
            with contextlib.suppress(NotFoundError, pywintypes.com_error):
                skip_ids.add(self.default_folder(f, mailbox).EntryID)
        found: list[tuple[str, Com]] = []

        def walk(parent: Com, path: str) -> None:
            for sub in parent.Folders:
                if sub.DefaultItemType != OlItemType.MAIL or sub.EntryID in skip_ids:
                    continue
                if sub.Name.lower() in SYSTEM_FOLDER_NAMES:
                    continue
                label = f"{path}/{sub.Name}" if path else sub.Name
                found.append((label, sub))
                walk(sub, label)

        walk(self.default_folder(OlFolder.INBOX, mailbox).Parent, "")
        return found

    def item(self, bid: BackendId) -> Com:
        """Item for a backend ID.

        Raises:
            NotFoundError: The item was moved or deleted.
        """
        try:
            return self.ns.GetItemFromID(bid.item_id, bid.container_id or self.default_store_id)
        except pywintypes.com_error as exc:
            raise NotFoundError("Item no longer exists (moved or deleted?). List again.") from exc

    def backend_id(self, item: Com, kind: str, store_id: str | None = None) -> BackendId:
        """Backend ID of an item; the mailbox (store) is taken from the item's folder unless given."""
        if store_id is None:
            try:
                store_id = str(item.Parent.StoreID)
            except (AttributeError, pywintypes.com_error):
                store_id = self.default_store_id
        return BackendId(kind=kind, item_id=item.EntryID, container_id=store_id)

    def set_sender(self, item: Com, mailbox: str) -> str:
        """Send `item` from another account in the profile, or on behalf of a shared mailbox.

        Returns a label of the sender for the preview.
        """
        account = next(
            (a for a in self.ns.Accounts if str(a.SmtpAddress).lower() == mailbox.strip().lower()), None
        )
        if account is not None:
            dispid = item._oleobj_.GetIDsOfNames("SendUsingAccount")
            item._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUTREF, False, account._oleobj_)
            return str(account.SmtpAddress)
        recipient = self.recipient(mailbox)
        item.SentOnBehalfOfName = self.smtp_of(recipient.AddressEntry) or recipient.Name
        return f"{recipient.Name} (on behalf)"

    def sender_for_store(self, store_id: str) -> str | None:
        """Mailbox to send from when replying to an item in `store_id` (None for the default mailbox)."""
        if store_id == self.default_store_id:
            return None
        store = next((s for s in self.ns.Stores if s.StoreID == store_id), None)
        return str(store.DisplayName) if store is not None else None

    @staticmethod
    def resolve_all(item: Com) -> None:
        """Resolve every recipient against the address book; refuse to continue if any fails."""
        recipients = item.Recipients
        if recipients.Count and not recipients.ResolveAll():
            raise UnresolvedRecipientsError([r.Name for r in recipients if not r.Resolved])

    def recipient(self, query: str) -> Com:
        """Resolve a name or email against the address book (same matching as typing it in Outlook)."""
        recipient = self.ns.CreateRecipient(query.strip())
        if not recipient.Resolve():
            raise UnresolvedRecipientsError([query])
        return recipient

    def smtp_of(self, address_entry: Com) -> str:
        """Primary SMTP address for an AddressEntry (Exchange users have X.500 addresses internally)."""
        if address_entry is None:
            return ""
        try:
            if address_entry.AddressEntryUserType in EXCHANGE_USER_TYPES:
                user = address_entry.GetExchangeUser()
                if user is not None:
                    return user.PrimarySmtpAddress
            return address_entry.Address or ""
        except pywintypes.com_error:
            return ""


class OutlookSession:
    """Thread-safe access to classic Outlook over COM.

    COM objects are bound to the (single-threaded apartment) thread that created them, while CLI and
    MCP code call us from arbitrary threads. The session therefore owns one dedicated worker thread;
    callers pass a function that receives an `Outlook` handle and runs *on that thread*. If Outlook
    stops responding (e.g. a hidden modal dialog), the call times out and the next call starts on a
    fresh thread and COM connection.

    Create once per process; `close()` on exit.
    """

    def __init__(self, *, timeout_s: float = 60.0) -> None:
        """Prepare the worker thread; Outlook is connected on the first call."""
        self._timeout_s = timeout_s
        self._lock = threading.Lock()
        self._executor = self._new_executor()
        self._local = threading.local()

    @staticmethod
    def _new_executor() -> cf.ThreadPoolExecutor:
        return cf.ThreadPoolExecutor(max_workers=1, thread_name_prefix="outlook-com", initializer=_co_init)

    def run[T](self, fn: Callable[[Outlook], T], *, timeout_s: float | None = None) -> T:
        """Run `fn(outlook)` on the COM thread and return its result.

        Raises:
            BackendUnavailableError: Outlook did not respond in time (the stuck thread is abandoned).
            NotFoundError, InvalidInputError: Outlook rejected the operation.
        """
        with self._lock:
            executor = self._executor
        future = executor.submit(self._invoke, fn)
        try:
            return future.result(timeout=timeout_s or self._timeout_s)
        except cf.TimeoutError as exc:
            with self._lock:
                if self._executor is executor:
                    executor.shutdown(wait=False, cancel_futures=True)
                    self._executor = self._new_executor()
            raise BackendUnavailableError(
                "Outlook did not respond in time (is a dialog open in classic Outlook?)."
            ) from exc

    def _invoke[T](self, fn: Callable[[Outlook], T]) -> T:
        outlook: Outlook | None = getattr(self._local, "outlook", None)
        if outlook is None:
            outlook = Outlook()
            self._local.outlook = outlook
        try:
            return fn(outlook)
        except M365Error:
            raise
        except pywintypes.com_error as exc:
            hresult, message = com_error_info(exc)
            if hresult == MAPI_E_NOT_FOUND:
                raise NotFoundError(message) from exc
            if "not connected" in message.lower():
                self._local.outlook = None
                raise BackendUnavailableError("Classic Outlook is not connected.") from exc
            raise InvalidInputError(f"Outlook rejected the operation: {message}") from exc

    def close(self) -> None:
        """Stop the worker thread."""
        self._executor.shutdown(wait=False, cancel_futures=True)


def _co_init() -> None:
    pythoncom.CoInitialize()
