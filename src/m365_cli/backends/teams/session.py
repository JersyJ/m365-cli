from __future__ import annotations

import json
import os
import subprocess
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any, Literal

from m365_cli.backends.teams.cdp import CdpConnection
from m365_cli.core.errors import BackendUnavailableError, M365Error, NotFoundError

TEAMS_URL = "https://teams.cloud.microsoft/"
SSO_STORE_URL = (
    "https://chromewebstore.google.com/detail/microsoft-single-sign-on/ppnbnpeolgkicgegkbkbjmhlideopiji"
)
TEAMS_HOSTS = ("teams.cloud.microsoft", "teams.microsoft.com", "teams.chat.microsoft.com")
LOGIN_HOSTS = ("login.microsoftonline.com", "login.microsoft.com", "login.live.com")
_HOOK_VERSION = 6
_START_WAIT = 25
_READY_WAIT = 60
_MAX_RETRIES = 3
_MAX_RETRY_WAIT = 20.0

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


@dataclass(frozen=True, slots=True)
class BrowserConfig:
    """How to find or start the dedicated Teams browser.

    Attributes:
        port: DevTools port on 127.0.0.1.
        profile_dir: Browser profile used only for Teams.
        executable: Chrome or Edge executable; None = find Chrome, then Edge.
    """

    port: int
    profile_dir: Path
    executable: Path | None = None


def _browser_candidates() -> list[Path]:
    roots = [os.environ.get(v, "") for v in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")]
    names = [r"Google\Chrome\Application\chrome.exe", r"Microsoft\Edge\Application\msedge.exe"]
    return [Path(root) / name for name in names for root in roots if root]


def _hook_source() -> str:
    return resources.files("m365_cli.backends.teams").joinpath("hook.js").read_text(encoding="utf-8")


class TeamsSession:
    """Calls the Teams web APIs from inside the signed-in Teams tab.

    The tokens never leave the browser: a small page script (hook.js) keeps the tokens that the Teams
    web app itself obtains, and runs our `fetch()` calls with them. The browser is a dedicated
    profile with a DevTools port on 127.0.0.1; it is started off-screen when needed and shown only
    when you must sign in.
    """

    def __init__(self, config: BrowserConfig) -> None:
        """Remember the browser settings; nothing starts before the first call."""
        self._config = config
        self._cdp: CdpConnection | None = None
        self._lock = threading.RLock()

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
                time.sleep(_retry_wait(reply.get("retryAfter"), attempt))
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
                import base64

                return base64.b64decode(str(reply.get("base64") or ""))
            _result(status, "", "GET", path)
            raise TeamsApiError(f"Teams answered {status} for the download.", status)
        raise TeamsAuthError("Teams did not provide a token for the download.")

    def status(self) -> dict[str, Any]:
        """Token and service state of the page script (no token values)."""
        with self._lock:
            return self._ready_cdp().evaluate("window.__m365cli.status()")

    def open_sso_install(self) -> None:
        """Open the Web Store page of the Microsoft Single Sign On extension in a visible window.

        Chrome does not allow a silent extension install (only an admin policy can force one),
        so the user adds it with one click on this page.
        """
        with self._lock:
            if self._tabs() is None:
                self._start_browser()
                self._wait_for_tabs()
            tab = self._open_tab(SSO_STORE_URL)
            cdp = CdpConnection(str(tab["webSocketDebuggerUrl"]))
            try:
                self._show_window(cdp)
            finally:
                cdp.close()

    def close(self) -> None:
        """Disconnect from the browser; the browser keeps running."""
        with self._lock:
            if self._cdp is not None:
                self._cdp.close()
                self._cdp = None

    def _call(self, payload: dict[str, Any]) -> dict[str, Any]:
        with self._lock:
            cdp = self._ready_cdp()
            reply = cdp.evaluate(f"window.__m365cli.call({json.dumps(payload)})", timeout=60)
        return reply if isinstance(reply, dict) else {"status": 0, "error": "no answer"}

    def _ready_cdp(self) -> CdpConnection:
        if self._cdp is None:
            self._cdp = self._connect()
        if not self._hook_ready(self._cdp):
            self._install_hook(self._cdp)
        return self._cdp

    def _renew(self) -> None:
        with self._lock:
            cdp = self._cdp or self._connect()
            self._cdp = cdp
            self._install_hook(cdp)

    def _hook_ready(self, cdp: CdpConnection) -> bool:
        state = cdp.evaluate(
            "window.__m365cli ? window.__m365cli.status() : null",
        )
        return (
            isinstance(state, dict)
            and state.get("version") == _HOOK_VERSION
            and bool(state.get("skype"))
            and bool(state.get("me"))
        )

    def _install_hook(self, cdp: CdpConnection) -> None:
        cdp.call("Page.enable")
        cdp.call("Page.addScriptToEvaluateOnNewDocument", {"source": _hook_source()})
        cdp.call("Page.navigate", {"url": TEAMS_URL})
        deadline = time.monotonic() + _READY_WAIT
        while time.monotonic() < deadline:
            time.sleep(1.5)
            try:
                href = str(cdp.evaluate("location.href") or "")
                if _host(href) in LOGIN_HOSTS:
                    self._show_window(cdp)
                    raise TeamsAuthError(
                        "Teams needs you to sign in. Sign in in the Teams browser window that is now "
                        "visible, then retry."
                    )
                if self._hook_ready(cdp):
                    return
            except TeamsAuthError:
                raise
            except BackendUnavailableError:
                continue
        self._show_window(cdp)
        raise TeamsAuthError(
            "Teams did not finish loading in the browser. Check the Teams window (now visible), sign in "
            "if asked, then retry."
        )

    def _connect(self) -> CdpConnection:
        tabs = self._tabs()
        if tabs is None:
            self._start_browser()
            tabs = self._wait_for_tabs()
        tab = next((t for t in tabs if _host(t.get("url", "")) in TEAMS_HOSTS), None)
        tab = tab or next((t for t in tabs if _host(t.get("url", "")) in LOGIN_HOSTS), None)
        if tab is None:
            tab = self._open_tab()
        return CdpConnection(str(tab["webSocketDebuggerUrl"]))

    def _endpoint(self, path: str) -> str:
        return f"http://127.0.0.1:{self._config.port}{path}"

    def _tabs(self) -> list[dict[str, Any]] | None:
        try:
            with urllib.request.urlopen(self._endpoint("/json/list"), timeout=3) as response:
                tabs = json.load(response)
        except (OSError, ValueError):
            return None
        return [t for t in tabs if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]

    def _wait_for_tabs(self) -> list[dict[str, Any]]:
        deadline = time.monotonic() + _START_WAIT
        while time.monotonic() < deadline:
            time.sleep(0.5)
            tabs = self._tabs()
            if tabs:
                return tabs
        raise BackendUnavailableError(
            f"The Teams browser did not open its DevTools port {self._config.port} (blocked by a policy, "
            "or the port is used by another program)."
        )

    def _open_tab(self, target_url: str = TEAMS_URL) -> dict[str, Any]:
        url = self._endpoint("/json/new?" + urllib.parse.quote(target_url, safe=""))
        request = urllib.request.Request(url, method="PUT")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return json.load(response)
        except (OSError, ValueError) as exc:
            raise BackendUnavailableError(f"Cannot open a Teams tab: {exc}") from exc

    def _start_browser(self) -> None:
        executable = self._config.executable or next((p for p in _browser_candidates() if p.exists()), None)
        if executable is None or not executable.exists():
            raise BackendUnavailableError("Chrome or Edge not found; set teams_browser in config.toml.")
        self._config.profile_dir.mkdir(parents=True, exist_ok=True)
        args = [
            str(executable),
            f"--remote-debugging-port={self._config.port}",
            "--remote-debugging-address=127.0.0.1",
            f"--user-data-dir={self._config.profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--window-position=-32000,-32000",
            TEAMS_URL,
        ]
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NEW_PROCESS_GROUP", 0
        )
        subprocess.Popen(
            args,
            creationflags=flags,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )

    def _show_window(self, cdp: CdpConnection) -> None:
        try:
            window = cdp.call("Browser.getWindowForTarget")
            bounds = {"left": 80, "top": 60, "width": 1280, "height": 860, "windowState": "normal"}
            cdp.call("Browser.setWindowBounds", {"windowId": window["windowId"], "bounds": bounds})
            cdp.call("Page.bringToFront")
        except M365Error:
            return


def _host(url: str) -> str:
    try:
        return urllib.parse.urlsplit(url).hostname or ""
    except ValueError:
        return ""


def _retry_wait(retry_after: object, attempt: int) -> float:
    try:
        seconds = float(str(retry_after))
    except ValueError:
        seconds = 2.0
    return min(max(seconds, 1.0) * (2**attempt), _MAX_RETRY_WAIT)


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
