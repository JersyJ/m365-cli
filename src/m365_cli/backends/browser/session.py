from __future__ import annotations

import json
import os
import shutil
import subprocess
import threading
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any, ClassVar

from m365_cli.backends.browser.cdp import CdpConnection, ConnectionLostError
from m365_cli.core.errors import BackendUnavailableError, M365Error

SSO_STORE_URL = (
    "https://chromewebstore.google.com/detail/microsoft-single-sign-on/ppnbnpeolgkicgegkbkbjmhlideopiji"
)
LOGIN_HOSTS = ("login.microsoftonline.com", "login.microsoft.com", "login.live.com")
HOOK_VERSION = 9
_START_WAIT = 25
_READY_WAIT = 60
_LINUX_BROWSERS = (
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "microsoft-edge",
    "microsoft-edge-stable",
)


@dataclass(frozen=True, slots=True)
class BrowserConfig:
    """How to find or start the dedicated browser that runs Teams and Outlook on the web.

    Attributes:
        port: DevTools port on 127.0.0.1.
        profile_dir: Browser profile used only by this tool.
        executable: Chrome or Edge executable; None = find Chrome, then Edge.
    """

    port: int
    profile_dir: Path
    executable: Path | None = None


def browser_candidates() -> list[Path]:
    """Chrome and Edge executables to try, Chrome first (Windows install folders, or the PATH elsewhere)."""
    if os.name != "nt":
        return [Path(found) for name in _LINUX_BROWSERS if (found := shutil.which(name))]
    roots = [os.environ.get(v, "") for v in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")]
    names = [r"Google\Chrome\Application\chrome.exe", r"Microsoft\Edge\Application\msedge.exe"]
    return [Path(root) / name for name in names for root in roots if root]


def _hook_source() -> str:
    return resources.files("m365_cli.backends.browser").joinpath("hook.js").read_text(encoding="utf-8")


def host_of(url: str) -> str:
    """Host name of a URL ('' when it has none)."""
    try:
        return urllib.parse.urlsplit(url).hostname or ""
    except ValueError:
        return ""


class PageSession:
    """One web app (Teams, Outlook) in its own tab of the dedicated browser.

    The tokens never leave the browser: a page script (hook.js) keeps the tokens that the web app
    itself obtains and runs our `fetch()` calls with them. The browser is a dedicated profile with a
    DevTools port on 127.0.0.1; it is started off-screen when needed and shown only for a sign-in.
    Subclasses name the app, its address and hosts, and when the page is ready.
    """

    app: ClassVar[str] = "the web app"
    url: ClassVar[str] = ""
    hosts: ClassVar[tuple[str, ...]] = ()
    auth_error: ClassVar[type[BackendUnavailableError]] = BackendUnavailableError

    def __init__(self, config: BrowserConfig) -> None:
        """Remember the browser settings; nothing starts before the first call."""
        self._config = config
        self._cdp: CdpConnection | None = None
        self._lock = threading.RLock()

    def _is_ready(self, state: dict[str, Any]) -> bool:
        """True when the page script has what the app needs (tokens, account)."""
        raise NotImplementedError

    def _recover(self, cdp: CdpConnection, href: str, waited: float) -> None:
        """Called while waiting for the page to become ready; apps override it to unstick a load."""

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
        """The tab's connection with the page script ready; a lost connection is replaced once."""
        ready = False
        if self._cdp is not None:
            try:
                ready = self._hook_ready(self._cdp)
            except ConnectionLostError:
                self._cdp = None
        if self._cdp is None:
            self._cdp = self._connect()
            ready = self._hook_ready(self._cdp)
        if not ready:
            self._install_hook(self._cdp)
        return self._cdp

    def _renew(self) -> None:
        with self._lock:
            cdp = self._cdp or self._connect()
            self._cdp = cdp
            self._install_hook(cdp)

    def _state(self, cdp: CdpConnection) -> dict[str, Any] | None:
        state = cdp.evaluate("window.__m365cli ? window.__m365cli.status() : null")
        return state if isinstance(state, dict) else None

    def _hook_ready(self, cdp: CdpConnection) -> bool:
        state = self._state(cdp)
        return state is not None and state.get("version") == HOOK_VERSION and self._is_ready(state)

    def _install_hook(self, cdp: CdpConnection) -> None:
        cdp.call("Page.enable")
        cdp.call("Page.addScriptToEvaluateOnNewDocument", {"source": _hook_source()})
        cdp.call("Page.navigate", {"url": self.url})
        started = time.monotonic()
        while time.monotonic() < started + _READY_WAIT:
            time.sleep(1.5)
            try:
                href = str(cdp.evaluate("location.href") or "")
                if host_of(href) in LOGIN_HOSTS:
                    self._show_window(cdp)
                    raise self.auth_error(
                        f"{self.app} needs you to sign in. Sign in in the browser window that is now "
                        "visible, then retry."
                    )
                if self._hook_ready(cdp):
                    return
                self._recover(cdp, href, time.monotonic() - started)
            except self.auth_error:
                raise
            except BackendUnavailableError:
                continue
        self._show_window(cdp)
        raise self.auth_error(
            f"{self.app} did not finish loading in the browser. Check the browser window (now visible), "
            "sign in if asked, then retry."
        )

    def _connect(self) -> CdpConnection:
        tabs = self._tabs()
        if tabs is None:
            self._start_browser()
            tabs = self._wait_for_tabs()
        tab = next((t for t in tabs if host_of(t.get("url", "")) in self.hosts), None)
        tab = tab or next((t for t in tabs if host_of(t.get("url", "")) in LOGIN_HOSTS), None)
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
            f"The browser did not open its DevTools port {self._config.port} (blocked by a policy, "
            "or the port is used by another program)."
        )

    def _open_tab(self, target_url: str | None = None) -> dict[str, Any]:
        url = self._endpoint("/json/new?" + urllib.parse.quote(target_url or self.url, safe=""))
        request = urllib.request.Request(url, method="PUT")
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return json.load(response)
        except (OSError, ValueError) as exc:
            raise BackendUnavailableError(f"Cannot open a {self.app} tab: {exc}") from exc

    def _start_browser(self) -> None:
        executable = self._config.executable or next((p for p in browser_candidates() if p.exists()), None)
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
            self.url,
        ]
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NEW_PROCESS_GROUP", 0
        )
        subprocess.Popen(
            args,
            creationflags=flags,
            start_new_session=os.name != "nt",
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


def retry_wait(retry_after: object, attempt: int, cap: float = 20.0) -> float:
    """Seconds to wait before retrying a 429: the server's Retry-After, doubled per attempt, capped."""
    try:
        seconds = float(str(retry_after))
    except ValueError:
        seconds = 2.0
    return min(max(seconds, 1.0) * (2**attempt), cap)
