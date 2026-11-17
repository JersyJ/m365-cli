from __future__ import annotations

import base64
import json
import time
from typing import Any, ClassVar, Literal

from m365_cli.backends.browser.cdp import CdpConnection
from m365_cli.backends.browser.session import PageSession, host_of, retry_wait
from m365_cli.core.errors import BackendUnavailableError, M365Error, NotFoundError

TEAMS_URL = "https://teams.cloud.microsoft/"
TEAMS_HOSTS = ("teams.cloud.microsoft", "teams.microsoft.com", "teams.chat.microsoft.com")
_NO_TOKEN_WAIT = 15
_MAX_RETRIES = 3

Auth = Literal["skype"] | str


class TeamsAuthError(BackendUnavailableError):
    """Teams refused the session (expired or missing sign-in)."""


class TeamsRateLimitError(M365Error):
    """Teams kept answering 429 (too many requests)."""


class TeamsApiError(M365Error):
    """Teams answered with an unexpected HTTP status."""

    def __init__(self, message: str, status: int) -> None:
        """Store the HTTP status next to the message."""
        super().__init__(message)
        self.status = status


class TeamsSession(PageSession):
    """Calls the Teams web APIs from inside the signed-in Teams tab of the dedicated browser."""

    app: ClassVar[str] = "Teams"
    url: ClassVar[str] = TEAMS_URL
    hosts: ClassVar[tuple[str, ...]] = TEAMS_HOSTS
    auth_error: ClassVar[type[BackendUnavailableError]] = TeamsAuthError
    _cleared = False

    def _is_ready(self, state: dict[str, Any]) -> bool:
        return bool(state.get("skype")) and bool(state.get("me"))

    def _install_hook(self, cdp: CdpConnection) -> None:
        self._cleared = False
        super()._install_hook(cdp)

    def _recover(self, cdp: CdpConnection, href: str, waited: float) -> None:
        """Teams reused a cached skype token and never called authz: clear its site data once."""
        if self._cleared or waited < _NO_TOKEN_WAIT:
            return
        state = self._state(cdp)
        if state is None or not state.get("me") or state.get("skype"):
            return
        self._cleared = True
        cdp.call(
            "Storage.clearDataForOrigin",
            {
                "origin": f"https://{host_of(href)}",
                "storageTypes": "local_storage,indexeddb,cache_storage,service_workers",
            },
        )
        cdp.call("Page.navigate", {"url": TEAMS_URL})

    def request(
        self,
        service: str,
        path: str,
        *,
        method: str = "GET",
        body: Any = None,
        auth: tuple[Auth, ...] = ("skype",),
        headers: dict[str, str] | None = None,
    ) -> Any:
        """Call one Teams service and return the parsed JSON body (None for an empty body).

        `service` is a key of the Teams service map (e.g. 'chatServiceAfd', 'chatSvcAggAfd') or
        'substrate' / 'presence'. `auth` lists 'skype' and/or bearer audiences.
        429 answers are retried with the server's Retry-After; one 401 reloads Teams to renew tokens.

        Raises:
            TeamsAuthError: Still unauthorized after renewing the tokens.
            TeamsRateLimitError: Still rate-limited after the retries.
            NotFoundError: 404.
            TeamsApiError: Any other unexpected status.
        """
        payload = {"service": service, "path": path, "method": method, "body": body, "auth": list(auth)}
        if headers:
            payload["headers"] = headers
        renewed = False
        for attempt in range(_MAX_RETRIES + 1):
            reply = self._call(payload)
            status = int(reply.get("status") or 0)
            if status == 0:
                error = str(reply.get("error", "no answer"))
                if error.startswith(("no-skype-token", "no-token")) and not renewed:
                    self._renew()
                    renewed = True
                    continue
                if error.startswith(("no-skype-token", "no-token")):
                    raise TeamsAuthError(
                        f"Teams has no token for {error.split(' ', 1)[-1]}. Open that part of Teams once "
                        "in the Teams window, then retry."
                    )
                raise BackendUnavailableError(f"Teams request failed: {error}")
            if status == 401 and not renewed:
                self._renew()
                renewed = True
                continue
            if status == 429 and attempt < _MAX_RETRIES:
                time.sleep(retry_wait(reply.get("retryAfter"), attempt))
                continue
            return _result(status, str(reply.get("body") or ""), method, path)
        raise TeamsRateLimitError("Teams is rate-limiting requests (429). Wait a minute, then retry.")

    def download(self, service: str, path: str, *, auth: tuple[Auth, ...]) -> bytes:
        """Download a binary file from a service (e.g. 'graph'); returns the raw bytes.

        Raises:
            TeamsAuthError: Unauthorized.
            NotFoundError: 404.
            TeamsApiError: Any other unexpected status.
        """
        payload = {"service": service, "path": path, "method": "GET", "auth": list(auth), "binary": True}
        renewed = False
        for _ in range(2):
            reply = self._call(payload)
            status = int(reply.get("status") or 0)
            error = str(reply.get("error", ""))
            if (status == 401 or error.startswith("no-token")) and not renewed:
                self._renew()
                renewed = True
                continue
            if status == 0:
                if error.startswith("no-token"):
                    raise TeamsAuthError(
                        f"Teams has no token for {error.split(' ', 1)[-1]}. Open a file in the Teams window "
                        "once, then retry."
                    )
                raise BackendUnavailableError(f"Teams download failed: {error or 'no answer'}")
            if 200 <= status < 300:
                return base64.b64decode(str(reply.get("base64") or ""))
            _result(status, "", "GET", path)
            raise TeamsApiError(f"Teams answered {status} for the download.", status)
        raise TeamsAuthError("Teams did not provide a token for the download.")


def _result(status: int, text: str, method: str, path: str) -> Any:
    if 200 <= status < 300:
        if not text.strip():
            return None
        try:
            return json.loads(text)
        except ValueError:
            return text
    where = f"{method} {path.split('?', 1)[0]}"
    if status in (401, 403):
        raise TeamsAuthError(f"Teams refused {where} ({status}). Your account may not have access.")
    if status == 404:
        raise NotFoundError(f"Teams item not found ({where}).")
    detail = text.strip().replace("\n", " ")[:200]
    raise TeamsApiError(f"Teams answered {status} for {where}: {detail}", status)
