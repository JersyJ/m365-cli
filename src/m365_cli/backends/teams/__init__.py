"""Teams backend: the Teams web APIs, called from inside the signed-in Teams tab of a dedicated browser.

The tokens stay in the browser. See `TeamsSession` and hook.js.
"""

from m365_cli.backends.teams.session import (
    BrowserConfig,
    TeamsApiError,
    TeamsAuthError,
    TeamsRateLimitError,
    TeamsSession,
)

__all__ = ["BrowserConfig", "TeamsApiError", "TeamsAuthError", "TeamsRateLimitError", "TeamsSession"]
