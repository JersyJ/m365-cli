#!/usr/bin/env bash
# One-command setup of m365-cli from WSL: installs Windows uv, clones the repo over SSH,
# builds the Windows venv, adds `m365`/`wuv` wrappers and registers the MCP server.
#
# Works standalone (a colleague gets just this file) or from inside an existing clone:
#   bash setup.sh
# With GitHub CLI access, without downloading the file first:
#   gh api repos/JersyJ/m365-cli/contents/scripts/setup.sh -H 'Accept: application/vnd.github.raw' | bash
#
# Needs: WSL with git, an SSH key added to GitHub with read access to the repo.
# Clones to C:\Users\<you>\m365-cli by default. Optional: M365_DIR=/mnt/c/Users/<you>/some/dir
# to clone elsewhere (anywhere under your Windows profile works).
# Safe to re-run. Later updates use scripts/update.sh.
set -euo pipefail

repo_url="git@github.com:JersyJ/m365-cli.git"
bindir="$HOME/.local/bin"
export GIT_SSH_COMMAND="${GIT_SSH_COMMAND:-ssh -o StrictHostKeyChecking=accept-new}"

step() { printf '\n==> %s\n' "$*"; }
die() { printf 'ERROR: %s\n' "$*" >&2; exit 1; }

# Everything runs inside main so `... | bash` reads the whole script before any command can eat stdin.
main() {
  # --- Windows side ------------------------------------------------------------
  command -v cmd.exe >/dev/null 2>&1 || die "Run this from WSL on Windows (cmd.exe not reachable)."
  command -v git >/dev/null 2>&1 || die "git is missing in WSL. Install it:  sudo apt install -y git"

  winprofile_raw="$(cd /mnt/c 2>/dev/null; cmd.exe /c 'echo %USERPROFILE%' 2>/dev/null | tr -d '\r')"
  winhome="$(wslpath -u "$winprofile_raw")"
  [[ -d "$winhome" ]] || die "Could not find your Windows profile folder ($winprofile_raw)."
  localappdata="$winhome/AppData/Local"

  # --- 1. Windows uv -----------------------------------------------------------
  find_uv() {
    local g
    for g in "$localappdata/Microsoft/WinGet/Packages/"astral-sh.uv_*/uv.exe "$winhome/.local/bin/uv.exe"; do
      [[ -f "$g" ]] && { echo "$g"; return 0; }
    done
    return 1
  }

  step "Windows uv"
  if uv_exe="$(find_uv)"; then
    echo "Found: $uv_exe"
  else
    echo "Not found, installing with winget ..."
    if command -v winget.exe >/dev/null 2>&1; then
      winget.exe install --id astral-sh.uv -e --silent \
        --accept-source-agreements --accept-package-agreements || true
    fi
    if ! uv_exe="$(find_uv)"; then
      echo "winget did not work, trying the official installer ..."
      powershell.exe -NoProfile -ExecutionPolicy Bypass -Command \
        "irm https://astral.sh/uv/install.ps1 | iex" || true
    fi
    uv_exe="$(find_uv)" || die "Could not install uv. In Windows run:  winget install astral-sh.uv  then re-run this script."
    echo "Installed: $uv_exe"
  fi
  chmod +x "$uv_exe" 2>/dev/null || true

  # --- 2. The repo ---------------------------------------------------------------
  step "Source code"
  here="$(cd "$(dirname "${BASH_SOURCE[0]:-.}")" && pwd)"
  if [[ -z "${M365_DIR:-}" && -n "${BASH_SOURCE[0]:-}" ]] && git -C "$here/.." remote get-url origin 2>/dev/null | grep -q 'm365-cli'; then
    repo="$(cd "$here/.." && pwd)"
    echo "Using this clone: $repo"
  else
    repo="${M365_DIR:-$winhome/m365-cli}"
    if [[ -d "$repo/.git" ]]; then
      echo "Already cloned at $repo, pulling ..."
      git -C "$repo" pull --ff-only
    else
      echo "Checking GitHub SSH access ..."
      ssh_out="$($GIT_SSH_COMMAND -T git@github.com 2>&1 </dev/null || true)"
      if [[ "$ssh_out" != *"successfully authenticated"* ]]; then
        echo "$ssh_out" >&2
        die "No SSH access to GitHub. Create a key (ssh-keygen -t ed25519), add the .pub file at
         https://github.com/settings/keys, ask for access to JersyJ/m365-cli, then re-run.
         If your key has a passphrase: eval \"\$(ssh-agent)\" && ssh-add"
      fi
      mkdir -p "$(dirname "$repo")"
      git clone "$repo_url" "$repo"
    fi
  fi

  # The venv is Windows Python (Outlook COM), so the repo must sit on a Windows drive.
  [[ "$repo" =~ ^/mnt/[a-z]/ ]] || die "The repo is at $repo. It must be on a Windows drive (/mnt/c/...)."

  # --- 3. Dependencies -------------------------------------------------------------
  step "Dependencies (builds the Windows venv, may download Python)"
  cd "$repo"
  if ! "$uv_exe" sync 2> >(tee /tmp/m365-setup.err >&2); then
    grep -q "being used by another process" /tmp/m365-setup.err \
      && die "The CLI binary is locked by a running Claude Code / MCP server. Close it and re-run."
    die "uv sync failed (see above)."
  fi
  m365_exe="$repo/.venv/Scripts/m365.exe"
  [[ -f "$m365_exe" ]] || die "uv sync finished but $m365_exe is missing."

  # --- 4. Wrappers -----------------------------------------------------------------
  step "Wrappers in $bindir"
  mkdir -p "$bindir"
  cat > "$bindir/wuv" <<EOF
#!/usr/bin/env bash
# Run the Windows uv.exe from WSL (cwd must be on a Windows drive).
exec "$uv_exe" "\$@"
EOF
  cat > "$bindir/m365" <<EOF
#!/usr/bin/env bash
exec "$m365_exe" "\$@"
EOF
  chmod +x "$bindir/wuv" "$bindir/m365"
  echo "Created wuv and m365."

  # --- 5. Register the MCP server ------------------------------------------------
  step "MCP server"
  registered=0
  if command -v claude >/dev/null 2>&1; then
    claude mcp remove m365 -s user >/dev/null 2>&1 || true
    claude mcp add m365 -s user -- "$m365_exe" mcp && registered=1 && echo "Registered with Claude Code."
  fi
  if command -v copilot >/dev/null 2>&1; then
    copilot mcp remove m365 >/dev/null 2>&1 || true
    copilot mcp add m365 -- "$m365_exe" mcp && registered=1 && echo "Registered with Copilot CLI."
  fi
  if [[ $registered -eq 0 ]]; then
    echo "Neither Claude Code nor Copilot CLI found in WSL. Register later with:"
    echo "  claude mcp add m365 -s user -- \"$m365_exe\" mcp"
  fi

  # --- Done ------------------------------------------------------------------------
  step "Done"
  case ":$PATH:" in
    *":$bindir:"*) : ;;
    *) echo "NOTE: $bindir is not on PATH. Add it, then reopen the shell:"
       echo "  echo 'export PATH=\"\$HOME/.local/bin:\$PATH\"' >> ~/.bashrc" ;;
  esac
  cat <<EOF
Next steps:
  1. Sign in to classic Outlook (once).
  2. Test it:     m365 mail list -n 5
  3. For Teams:   m365 teams browser   (signs in via the SSO extension)
  4. Restart Claude Code / Copilot so it loads the MCP server.

Update later:     bash $repo/scripts/update.sh
EOF
}

main "$@"
