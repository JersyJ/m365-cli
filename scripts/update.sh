#!/usr/bin/env bash
# Update m365-cli: pull the latest source and sync dependencies. Run from WSL.
# Source changes apply when Claude Code restarts (the project is an editable install).
# `uv sync` rebuilds the CLI binary, which is locked while Claude Code / the MCP server runs,
# so if dependencies changed, close Claude Code first.
set -euo pipefail
repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo"
echo "Updating $repo ..."
git pull --ff-only
uv=uv
command -v wuv >/dev/null 2>&1 && uv=wuv
if "$uv" sync 2>/tmp/m365-update.err; then
  echo "Done. Restart Claude Code so the MCP server reloads the new version."
elif grep -q "being used by another process" /tmp/m365-update.err; then
  echo "Source updated (applies on restart). The CLI binary is locked by a running Claude Code."
  echo "To also update dependencies: close Claude Code, then run  $uv sync"
else
  cat /tmp/m365-update.err >&2
  echo "Dependency sync failed (see above). If it is an SSH/auth error, run: ssh-add ~/.ssh/personal" >&2
  exit 1
fi
