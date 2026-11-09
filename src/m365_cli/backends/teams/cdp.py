from __future__ import annotations

import json
import threading
import time
from typing import Any

from websockets.exceptions import WebSocketException
from websockets.sync.client import ClientConnection, connect

from m365_cli.core.errors import BackendUnavailableError

_MAX_MESSAGE = 64 * 1024 * 1024


class CdpError(BackendUnavailableError):
    """The browser returned an error for a DevTools call, or the page script threw."""


class CdpConnection:
    """One DevTools websocket to a browser tab; calls are serialized, events are ignored."""

    def __init__(self, ws_url: str, *, timeout: float = 10) -> None:
        """Connect to the tab's `webSocketDebuggerUrl`.

        Raises:
            BackendUnavailableError: The browser does not accept the connection.
        """
        try:
            self._ws: ClientConnection = connect(
                ws_url, open_timeout=timeout, max_size=_MAX_MESSAGE, origin=None, compression=None
            )
        except (OSError, WebSocketException) as exc:
            raise BackendUnavailableError(f"Cannot connect to the Teams browser tab: {exc}") from exc
        self._next_id = 0
        self._lock = threading.Lock()

    def call(
        self, method: str, params: dict[str, Any] | None = None, *, timeout: float = 30
    ) -> dict[str, Any]:
        """Send one DevTools command and wait for its result.

        Raises:
            CdpError: The browser reports an error.
            BackendUnavailableError: The connection is lost or the call times out.
        """
        with self._lock:
            self._next_id += 1
            call_id = self._next_id
            deadline = time.monotonic() + timeout
            try:
                self._ws.send(json.dumps({"id": call_id, "method": method, "params": params or {}}))
                while True:
                    left = deadline - time.monotonic()
                    if left <= 0:
                        raise BackendUnavailableError(
                            f"The Teams tab did not answer {method} in {timeout:g} s."
                        )
                    message = json.loads(self._ws.recv(timeout=left))
                    if message.get("id") == call_id:
                        break
            except TimeoutError as exc:
                raise BackendUnavailableError(
                    f"The Teams tab did not answer {method} in {timeout:g} s."
                ) from exc
            except (OSError, WebSocketException) as exc:
                raise BackendUnavailableError(f"Lost the connection to the Teams tab: {exc}") from exc
        if "error" in message:
            raise CdpError(f"{method}: {message['error'].get('message', message['error'])}")
        return message.get("result", {})

    def evaluate(self, expression: str, *, timeout: float = 30) -> Any:
        """Evaluate JavaScript in the page (promises are awaited) and return the JSON value.

        Raises:
            CdpError: The script threw.
        """
        result = self.call(
            "Runtime.evaluate",
            {"expression": expression, "awaitPromise": True, "returnByValue": True},
            timeout=timeout,
        )
        if "exceptionDetails" in result:
            details = result["exceptionDetails"]
            text = details.get("exception", {}).get("description") or details.get("text", "error")
            raise CdpError(f"Script error in the Teams tab: {text.splitlines()[0]}")
        return result.get("result", {}).get("value")

    def close(self) -> None:
        """Close the websocket."""
        self._ws.close()
