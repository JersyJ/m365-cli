"""Briefing feature: a cross-feature 'what needs me today' summary of mail, calendar and Teams.

This is an aggregator: unlike other features it calls the public APIs of mail, calendar and Teams.
Each section is skipped when its feature is disabled and degrades gracefully when a backend is down.
"""

from m365_cli.core.feature import Feature
from m365_cli.features.briefing.cli import build_cli
from m365_cli.features.briefing.mcp import register_mcp

FEATURE = Feature(
    name="briefing",
    help="A daily 'what needs me' summary across mail, calendar and Teams.",
    build_cli=build_cli,
    register_mcp=register_mcp,
)
