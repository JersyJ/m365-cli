# m365-cli

m365-cli gives an AI agent access to your own Outlook mailbox, calendar and Teams chats.
It has an MCP server for Claude Code (`m365 mcp`) and a CLI (`m365`). The results are short text, so the agent uses few tokens.

The tool controls classic Outlook through COM and uses the Outlook profile that is already signed in.
For Teams, it uses the Teams web app in a separate browser profile. The calls run inside the signed-in Teams tab, so the sign-in tokens stay in the browser.
It needs no Graph app registration and no admin consent.

## Tools

Each MCP tool has a CLI command with the same function.

### Mail

| MCP tool | CLI | What it does |
|---|---|---|
| `mail_find` | `m365 mail list` | Finds mail in a folder or in all folders (`folder="all"`). Filters: sender, subject, text, date, unread, flagged, attachments. Pages with `offset`. |
| `mail_digest` | `m365 mail digest` | Groups many mails by sender: count, unread, newest time, top subjects. Use it for "what is new". |
| `mail_read` | `m365 mail read` | Shows one mail as plain text, without the quoted history (`detailed` shows all). |
| `mail_thread` | `m365 mail thread` | Shows a full conversation over all folders, from the first message to the last. |
| `mail_waiting` | `m365 mail waiting` | Shows sent mails that have no reply after N days. |
| `mail_mailboxes` | `m365 mail mailboxes` | Shows the mailboxes you can use: your mailbox, other accounts, shared mailboxes. |
| `mail_draft` | `m365 mail new` / `reply` / `forward` | Saves a draft: new, reply, reply-all or forward. Adds your Outlook font and signature. The body is Markdown. |
| `mail_send` | `m365 mail send` | Sends a draft now, or at a later time (`send_at`). |
| `mail_update` | `m365 mail update` | Marks read or unread, flags, sets categories, moves, deletes. Works on one mail, many refs, or all mails that agree with a filter. |
| `mail_attachments` | `m365 mail attachments` | Lists or saves the attachments, or shows the text of a PDF, DOCX, PPTX, XLSX or text file. |

Most mail tools accept `mailbox` (a shared mailbox or another account). `mail_draft` accepts `from_mailbox`.
A reply to a mail in another mailbox goes out from that mailbox.

### Calendar

| MCP tool | CLI | What it does |
|---|---|---|
| `cal_find` | `m365 calendar list` | Lists the events of a period, also recurring and moved occurrences. With `owner`, shows the calendar of a colleague who shares it with you. |
| `cal_read` | `m365 calendar read` | Shows one event: organizer, attendees, Teams link, description. |
| `cal_find_free_time` | `m365 calendar free` | Finds periods in working hours when you and the attendees are free. |
| `cal_draft` | `m365 calendar new` | Saves a meeting (not sent) or a personal appointment. `repeat` makes a recurring meeting. |
| `cal_update` | `m365 calendar update` | Moves a meeting, or changes attendees, title or location. The attendees get no message yet. |
| `cal_send` | `m365 calendar send` | Sends the invitation or the update of a meeting that you organize. |
| `cal_respond` | `m365 calendar respond` | Accepts, tentatively accepts or declines a meeting. |
| `cal_cancel` | `m365 calendar cancel` | Cancels a meeting that you organize, or deletes a personal appointment. |
| `cal_cleanup` | `m365 calendar cleanup` | Removes canceled meetings of other people from your calendar. It notifies nobody. |

### Briefing

| MCP tool | CLI | What it does |
|---|---|---|
| `briefing` | `m365 briefing` | One "what needs me today" summary across calendar, mail and Teams: today's meetings, unread and awaiting mail, and Teams mentions and unread. |

### People

| MCP tool | CLI | What it does |
|---|---|---|
| `people_lookup` | `m365 people find` / `status` | Shows the contact card (email, title, department, phone, manager) and the availability: free or busy now, out of office. |

### Teams

| MCP tool | CLI | What it does |
|---|---|---|
| `teams_find` | `m365 teams list` | Lists chats, group chats, meeting chats and channels, newest activity first. Filters: name, kind, unread. |
| `teams_read` | `m365 teams read` | Shows the newest messages of a chat: sender, time, text, quoted message, reactions, files. Pages with `older`. |
| `teams_digest` | `m365 teams digest` | What needs you: @mentions and replies to you, plus unread chats and channels. Use it for "what is new". |
| `teams_search` | `m365 teams search` | Full-text search across your Teams chats and channels. Each hit links to its conversation. |
| `teams_file` | `m365 teams files` | Lists files shared in a chat, or reads the text of one (PDF, DOCX, PPTX, XLSX, and text files such as .txt, .md, .csv, .json). |
| `teams_draft` | `m365 teams write` | Prepares a message, a reply (with quote) or a first message to a person. The body is Markdown. `@[Name]` mentions a chat member. |
| `teams_send` | `m365 teams send` | Sends a draft now, or at a later time (`send_at`). |
| `teams_status` | `m365 teams status` | Shows a colleague's live Teams presence (available, busy, in a meeting, away, out of office, offline), status note and card. |
| `teams_mark_read` | `m365 teams mark-read` | Marks a chat or channel read up to its newest message (only your read state). |
| `teams_update` | `m365 teams react` / `edit` / `delete` | Reacts to a message, or edits/deletes one of your own. Visible to others, so it needs approval. |

A chat can be a ref (`c:…`) or a name: the chat title, or the other person of a 1:1 chat.
For a search across both mail and Teams, the agent combines `mail_find` and `teams_search`.

### Limits

* A mail with rights protection (Microsoft Information Protection) or S/MIME has the mark `[protected]`. The tool cannot read its content.
* Most colleagues share only free/busy. Then `cal_find` with `owner` gives an error and tells you to use free/busy.
* Teams uses your home account only. Guest accounts in other tenants are not used.
* Teams has no supported API for this. The tool uses the internal API of the Teams web app, which can change. See [docs/teams-api.md](docs/teams-api.md).
* The first Teams listing can take one minute, because the tool looks up the names of 1:1 chats. Later listings use a cache.

## Safety

* Each send has two steps. `mail_draft`, `cal_draft`, `cal_update` and `teams_draft` save and show a preview. Only `mail_send`, `cal_send`, `cal_respond`, `cal_cancel` and `teams_send` send.
* The five send tools have the flag `anthropic/requiresUserInteraction`. Claude Code asks you before each call. Allow rules and auto mode do not skip the question.
* A Teams draft stays on your computer until `teams_send`. The tool never sends to the chats in `teams_protected`.
* A send limit applies (default: 5 per minute). All sends go into a log. Run `m365 log` to see the log.
* Mail bodies, attachment text, event descriptions and Teams messages are between `<<<UNTRUSTED …>>>` markers. The server instructions tell the model to not obey text from other people.
* The tool removes control characters and invisible characters from all output.
* The CLI asks before it sends. Use `--yes` to send without a question. Use `--draft` to save only.

## Install (Windows)

### What you need

* **Classic Outlook**, signed in. It can run in the background. Used for mail, calendar and people.
* **Google Chrome** ([google.com/chrome](https://www.google.com/chrome/)), only if you want Teams. The tool drives Teams in a separate Chrome profile.
* The **Microsoft Single Sign On** Chrome extension ([web store](https://chromewebstore.google.com/detail/microsoft-single-sign-on/ppnbnpeolgkicgegkbkbjmhlideopiji)), only for Teams. On a work-joined (Entra/Intune) PC it is usually installed for you, and it signs Teams in automatically with your Windows login — no password prompt.

### First-time install

1. Install **uv** and the **GitHub CLI** (the GitHub CLI is needed only while this repo is private):

   ```powershell
   winget install astral-sh.uv GitHub.cli
   gh auth login
   ```

2. Download the latest release and install it as a tool. This puts the `m365` command on your PATH (`%USERPROFILE%\.local\bin\m365.exe`) with all dependencies:

   ```powershell
   gh release download --repo JersyJ/m365-cli --pattern '*.whl'
   uv tool install m365_cli-*.whl
   ```

3. Sign in to classic Outlook (once), then test it:

   ```powershell
   m365 mail list -n 5
   ```

4. Teams (optional): run `m365 teams browser` once to check the sign-in. The tool uses its own Chrome profile in `%LOCALAPPDATA%\m365-cli\chrome-teams`, started off-screen. On a work-joined (Intune) PC the Single Sign On extension is force-installed by policy into every profile, so Teams signs in by itself — the output shows `"skype"` and `"me"` when it is ready. Without that policy the fresh profile has no extension: the tool shows the window, and you sign in once with your work account.

   Later Teams commands start the browser by themselves. Keep it open if you use Teams often: a cold start takes about 20 seconds.

5. Connect your agent — see the next sections.

### Update

```powershell
m365 update
```

It checks the latest release, downloads the wheel and reinstalls it in the background (a few seconds). Check with `m365 version`. The token comes from `gh` or the `GH_TOKEN` variable; once the repo is public, no token is needed.

### For developers

Work from the source instead of a release: clone the repo anywhere under `C:\Users\<you>\`, run `uv sync`, and use `uv run m365 …` (or `.venv\Scripts\m365.exe`). Update with `git pull` + `uv sync`; close Claude Code first if dependencies changed, because it locks the CLI binary. `m365 update` is only for the release install.

To cut a release: bump the version in `pyproject.toml` (the only place — the code reads it from the package metadata), commit, and push a matching tag, for example `git tag -m v0.2.0 v0.2.0 && git push origin main v0.2.0`. The workflow in `.github/workflows/release.yml` builds the wheel and attaches it to a GitHub Release; `m365 update` picks it up from there.

### Connect Claude Code

Run this command once (in WSL or in Windows), then start a new Claude Code session:

```bash
claude mcp add m365 -s user -- /mnt/c/Users/<you>/.local/bin/m365.exe mcp
```

### Connect GitHub Copilot

Copilot CLI (in WSL or in Windows):

```bash
copilot mcp add m365 -- /mnt/c/Users/<you>/.local/bin/m365.exe mcp
```

VS Code (Copilot agent mode): add this server to the user file `%APPDATA%\Code\User\mcp.json`, then select "Agent" in Copilot Chat:

```json
{
  "servers": {
    "m365": {
      "type": "stdio",
      "command": "C:\\Users\\<you>\\.local\\bin\\m365.exe",
      "args": ["mcp"]
    }
  }
}
```

Safety with Copilot: the flag `anthropic/requiresUserInteraction` works only in Claude Code.
Copilot asks before each tool call by default. Do not auto-approve the send tools.
In Copilot CLI, block them for unattended runs: `--deny-tool 'm365(mail_send)' --deny-tool 'm365(cal_send)' --deny-tool 'm365(cal_respond)' --deny-tool 'm365(cal_cancel)' --deny-tool 'm365(teams_send)'`.

## Settings

The file `%APPDATA%\m365-cli\config.toml` is optional. All keys are optional.

```toml
disabled_features = ["people"]
signature_new = "auto"
signature_reply = "auto"
signoff = ""
list_limit = 15
body_max_chars = 6000
work_hours = [8, 17]
max_sends_per_minute = 5
teams_protected = ["All Company*", "*Announcements*"]
teams_port = 9222
```

| Key | Meaning |
|---|---|
| `disabled_features` | Features to remove (`mail`, `calendar`, `people`, `teams`, `briefing`). Fewer tools use fewer tokens. |
| `signature_new`, `signature_reply` | `"auto"` (the Outlook default), `"none"`, or the name of an Outlook signature. |
| `signoff` | Text that the tool adds to each text that it sends. |
| `list_limit` | Number of items in a list (default 15). |
| `body_max_chars` | Maximum length of a mail body in the result (default 6000). |
| `work_hours` | Start hour and end hour for free-time searches. |
| `max_sends_per_minute` | Send limit for mail, invitations and Teams messages (default 5). |
| `teams_protected` | Teams chat titles that the tool never sends to. `*` is a wildcard. |
| `teams_port` | DevTools port of the Teams browser on 127.0.0.1 (default 9222). |
| `teams_browser`, `teams_profile_dir` | Browser executable (default: Chrome, then Edge) and its profile directory. |

## More documentation

* [Architecture and code rules](docs/architecture.md)
* [Teams web API notes](docs/teams-api.md)
