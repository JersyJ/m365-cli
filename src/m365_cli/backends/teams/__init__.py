"""Teams backend: the Teams web APIs, called from inside the signed-in Teams tab of the dedicated browser.

The tokens stay in the browser. See `TeamsSession` and hook.js.
"""

from m365_cli.backends.browser import BrowserConfig
from m365_cli.backends.teams.session import (
    TeamsApiError,
    TeamsAuthError,
    TeamsRateLimitError,
    TeamsSession,
)

__all__ = ["BrowserConfig", "TeamsApiError", "TeamsAuthError", "TeamsRateLimitError", "TeamsSession"]
