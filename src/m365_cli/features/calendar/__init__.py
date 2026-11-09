"""Calendar feature: events, invitations (draft/send/update/cancel), responses, common free time."""

from m365_cli.core.feature import Feature
from m365_cli.features.calendar.cli import build_cli
from m365_cli.features.calendar.mcp import INSTRUCTIONS, register_mcp

FEATURE = Feature(
    name="calendar",
    help="Events, invitations, responses and free time.",
    build_cli=build_cli,
    register_mcp=register_mcp,
    instructions=INSTRUCTIONS,
)
