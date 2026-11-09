"""Classic Outlook (COM) backend. Uses the signed-in Outlook profile; no Graph, no app registration."""

from m365_cli.backends.outlook.session import Com, Outlook, OutlookSession, iterate

__all__ = ["Com", "Outlook", "OutlookSession", "iterate"]
