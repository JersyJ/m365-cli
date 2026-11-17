"""Dedicated browser for the web apps (Teams, Outlook): DevTools connection, page script, one session per tab.

The tokens stay in the browser. See `PageSession` and hook.js.
"""

from m365_cli.backends.browser.session import BrowserConfig, PageSession

__all__ = ["BrowserConfig", "PageSession"]
