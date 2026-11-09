from __future__ import annotations

import importlib
import pkgutil
from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import typer
    from mcp.server.mcpserver import MCPServer

    from m365_cli.core.config import Settings
    from m365_cli.core.context import AppContext

ContextProvider = Callable[[], "AppContext"]


@dataclass(frozen=True, slots=True)
class Feature:
    """Feature plug-in contract.

    A feature is a self-contained package under `m365_cli.features` that exposes
    `FEATURE = Feature(...)`. It contributes a Typer sub-app (CLI) and/or MCP tools; the front-ends
    only iterate over enabled features, so adding or removing one never touches the rest of the code.

    Attributes:
        name: CLI sub-command name and key in `disabled_features`.
        help: One-line summary.
        build_cli: Builds the Typer sub-app (optional).
        register_mcp: Adds the MCP tools to a server (optional).
        instructions: Appended to the MCP server instructions when the feature is enabled.
    """

    name: str
    help: str
    build_cli: Callable[[ContextProvider], typer.Typer] | None = None
    register_mcp: Callable[[MCPServer, ContextProvider], None] | None = None
    instructions: str = ""


def discover_features(settings: Settings) -> list[Feature]:
    """Import every `m365_cli.features.<name>` package and return its FEATURE unless disabled."""
    import m365_cli.features as pkg

    disabled = set(settings.disabled_features)
    found: list[Feature] = []
    for info in sorted(pkgutil.iter_modules(pkg.__path__), key=lambda i: i.name):
        if not info.ispkg or info.name in disabled:
            continue
        module = importlib.import_module(f"{pkg.__name__}.{info.name}")
        feature = getattr(module, "FEATURE", None)
        if isinstance(feature, Feature):
            found.append(feature)
    return found
