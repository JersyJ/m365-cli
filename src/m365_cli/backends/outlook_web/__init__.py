"""Outlook on the web backend: the Outlook REST API and service.svc, called from inside the Outlook tab.

The tab runs in the same dedicated browser as Teams; the tokens stay in the browser. See
`OutlookWebSession` and `m365_cli.backends.browser`.
"""

from m365_cli.backends.outlook_web.session import (
    OutlookApiError,
    OutlookAuthError,
    OutlookWebSession,
)

__all__ = ["OutlookApiError", "OutlookAuthError", "OutlookWebSession"]
