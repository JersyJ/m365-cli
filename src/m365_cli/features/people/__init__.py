"""People feature: contact cards and availability / out-of-office from the Exchange address book."""

from m365_cli.core.feature import Feature
from m365_cli.features.people.cli import build_cli
from m365_cli.features.people.mcp import register_mcp

FEATURE = Feature(
    name="people",
    help="Look up colleagues and their availability / out of office.",
    build_cli=build_cli,
    register_mcp=register_mcp,
)
