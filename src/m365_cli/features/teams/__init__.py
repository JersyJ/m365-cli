"""Teams feature: find and read chats, then draft and send messages and replies.

The real backend calls the Teams web APIs from inside the signed-in Teams tab of a dedicated browser
(see `m365_cli.backends.teams`); the tokens never leave the browser.
"""

from m365_cli.core.feature import Feature
from m365_cli.features.teams.cli import build_cli
from m365_cli.features.teams.mcp import INSTRUCTIONS, register_mcp

FEATURE = Feature(
    name="teams",
    help="Teams chats, channels and presence: find, read, draft, send, status.",
    build_cli=build_cli,
    register_mcp=register_mcp,
    instructions=INSTRUCTIONS,
)
