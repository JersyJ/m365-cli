# Architecture and code rules

## Structure

```
src/m365_cli/
  cli.py, mcp_server.py   front-ends; they use only the enabled features
  core/                   shared parts: settings, errors, refs, paging, send log, text, markup and time helpers
  backends/outlook/       COM session (one STA thread, time-out), constants, free/busy, profile (signatures, fonts)
  backends/browser/       the dedicated browser: DevTools connection, page script (hook.js), one session per tab
  backends/teams/         Teams session (a tab of the browser)
  backends/outlook_web/   Outlook on the web session (a tab of the browser): REST API and service.svc
  backends/fake/          fake backend (JSON state), test data, snapshot of real data
  features/<name>/        api.py, service.py (classic Outlook or Teams), web.py (Outlook on the web), fake.py,
                          models.py, rules.py, render.py, cli.py, mcp.py
```

## Rules

* A feature uses `core` and `backends`. A feature never uses another feature, except the `briefing` aggregator, which calls the public `*_api` of mail, calendar and Teams and degrades gracefully when one is disabled or unavailable.
* `api.py` defines the interface. `service.py` (classic Outlook or Teams), `web.py` (Outlook on the web) and `fake.py` (fake backend) supply it.
* `rules.py` has the logic that all backends share. `render.py` makes the short text.
* `cli.py` and `mcp.py` are thin adapters. They do not contain logic.
* A backend starts at its first use. A command that does not use Outlook does not start Outlook.

## The browser backends (Teams, Outlook on the web)

The tool starts a separate browser profile with a DevTools port on 127.0.0.1. Teams and Outlook on the web each have one tab.
A page script (`hook.js`) keeps the tokens that the web app gets for your home account.
`PageSession` runs a `fetch()` inside the tab, so the tokens never leave the browser.
In Teams, the tool calls the service addresses from the service map of your account; in Outlook, only the tab's own origin.

Outlook on the web: the REST API (`/api/beta`) for most calls, `service.svc` where REST notifies people or cannot do it.
Refs hold immutable ids and the mailbox, so they survive moves and sending.
A meeting is created through REST without attendees; attendees and later changes go through `service.svc` with `SendToNone`.
A REST change to a meeting with attendees sends an update at once, so the tool never makes one.

## The fake backend

The fake backend acts like classic Outlook, but keeps all data in one JSON file and never sends.
The unit tests and the evaluations use it. It is in the package because the evaluations start the real MCP server
(`m365 mcp --fake-state …`). The fake backend loads only with `--fake-state`.

## Add a feature

1. Make the directory `src/m365_cli/features/<name>/` with the files of the structure above.
2. In `__init__.py`, export `FEATURE = Feature(name=..., help=..., build_cli=..., register_mcp=...)`.
3. Do nothing more. The front-ends find the feature automatically.

To remove a feature, add its name to `disabled_features`, or delete its directory.

## Rules for code text

* Do not write `#` comments. Write docstrings.
* Only an `__init__.py` file has a module docstring.
* ruff (`D` rules, Google convention) and pyrefly check the code.
