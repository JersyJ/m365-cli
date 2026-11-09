"""Mail feature: find, read, draft/send (new, reply, reply-all, forward), flag/move/delete, attachments."""

from m365_cli.core.feature import Feature
from m365_cli.features.mail.cli import build_cli
from m365_cli.features.mail.mcp import INSTRUCTIONS, register_mcp

FEATURE = Feature(
    name="mail",
    help="Read, search, compose and manage mail.",
    build_cli=build_cli,
    register_mcp=register_mcp,
    instructions=INSTRUCTIONS,
)
