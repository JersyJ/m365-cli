# m365-cli

m365-cli gives an AI agent access to your own Microsoft 365 mailbox and calendar.
It has two front-ends: a CLI (`m365`) and an MCP server (`m365 mcp`).
The results are short text, so the agent uses few tokens.

The tool controls classic Outlook through COM. It uses the Outlook profile that is already signed in.
Thus, it needs no Graph app registration and no admin consent. The Intune device check stays valid.
Teams support (through Teams on the web in Chrome) is the next feature.

## Features

| Feature    | CLI               | MCP tools |
|------------|-------------------|-----------|
| `mail`     | `m365 mail …`     | `mail_find`, `mail_read`, `mail_draft`, `mail_send`, `mail_update`, `mail_attachments` |
| `calendar` | `m365 calendar …` | `cal_find`, `cal_read`, `cal_draft`, `cal_send`, `cal_respond`, `cal_cancel`, `cal_find_free_time` |
| `people`   | `m365 people …`   | `people_lookup` |

Mail: find, read, draft/send (new, reply, reply-all, forward), flag/move/delete, attachments.
Calendar: events (also recurring and moved occurrences), invitations, responses, common free time.
People: contact card (email, title, department, phone, manager), availability and out of office.

## Safety

* Each send has two steps. `*_draft` saves a draft and shows a preview. Only `mail_send`, `cal_send`, `cal_respond` and `cal_cancel` send.
* The four send tools have the flag `anthropic/requiresUserInteraction`. Claude Code asks you before each call. Allow rules and auto mode do not skip the question.
* A send limit applies (default: 5 per minute). All sends go into a log. Run `m365 log` to see the log.
* Mail bodies and event descriptions are between `<<<UNTRUSTED …>>>` markers. The server instructions tell the model to not obey text from other people.
* The tool removes control characters and invisible characters from all output.
* The CLI asks before it sends. Use `--yes` to send without a question. Use `--draft` to save only.

## Install (Windows)

1. Install uv: `winget install astral-sh.uv`
2. Open the project directory: `cd C:\Users\<you>\src\m365-cli`
3. Install the dependencies: `uv sync`
4. Open classic Outlook and sign in. Outlook can run in the background.
5. Do a test: `uv run m365 mail list -n 5`

In WSL, call `/mnt/c/Users/<you>/src/m365-cli/.venv/Scripts/m365.exe`.
A wrapper in `~/.local/bin/m365` makes this easier.

## Connect Claude Code

Run this command once (in WSL or in Windows):

```bash
claude mcp add m365 -s user -- /mnt/c/Users/<you>/src/m365-cli/.venv/Scripts/m365.exe mcp
```

Then start a new Claude Code session.

## Settings

The file `%APPDATA%\m365-cli\config.toml` is optional. All keys are optional.

```toml
disabled_features = ["people"]
list_limit = 15
body_max_chars = 6000
signoff = ""
work_hours = [8, 17]
max_sends_per_minute = 5
```

* `disabled_features` removes features. Fewer tools use fewer tokens.
* `signoff` is text that the tool adds to each sent text.

## Architecture

```
src/m365_cli/
  cli.py, mcp_server.py   front-ends; they use only the enabled features
  core/                   shared parts: settings, errors, refs, paging, send log, text and time helpers
  backends/outlook/       COM session (one STA thread, time-out), constants, free/busy
  backends/fake/          fake backend (JSON state), test data, snapshot of real data
  features/<name>/        api.py, service.py (Outlook), fake.py, models.py, rules.py, render.py, cli.py, mcp.py
```

Rules:

* A feature uses `core` and `backends`. A feature never uses another feature.
* `api.py` defines the interface. `service.py` (Outlook) and `fake.py` (fake backend) supply it.
* `rules.py` has the logic that all backends share. `render.py` makes the short text.
* `cli.py` and `mcp.py` are thin adapters. They do not contain logic.
* A backend starts at its first use. A command that does not use Outlook does not start Outlook.

### Add a feature

1. Make the directory `src/m365_cli/features/<name>/` with the files of the table above.
2. In `__init__.py`, export `FEATURE = Feature(name=..., help=..., build_cli=..., register_mcp=...)`.
3. Do nothing more. The front-ends find the feature automatically.

To remove a feature, add its name to `disabled_features`, or delete its directory.

## Outlook quirks

* pywin32 labels Outlook times as UTC, but the values are local times. pywin32 also shifts naive times when it writes. `core.timeutil.from_com` and `to_com` correct this.
* Jet filters read dates with the Windows locale. On Czech Windows, Jet reads `2026-11-09` as 11 September. Thus, all filters use DASL with UTC ISO dates.
* Occurrences of a recurring meeting have the same EntryID as the series. Thus, the ref of an occurrence also contains its start time.
* `Recipient.FreeBusy` can return no data for your own mailbox. Thus, the tool reads your own calendar.
* All COM calls run on one thread with a time-out. A stuck call does not stop the next call.

## Tests and evaluations

Run the checks:

```powershell
uv run ruff format; uv run ruff check; uv run pyright; uv run pytest
```

The unit tests use the fake backend with synthetic data. They do not need Outlook.

The evaluations run real Claude Code sessions on the fake backend. See `evals/README.md`.

## Rules for code text

* Do not write `#` comments. Write docstrings.
* Only an `__init__.py` file has a module docstring.
* ruff (`D` rules, Google convention) and pyright check the code.
