# Architecture and code rules

## Structure

```
src/m365_cli/
  cli.py, mcp_server.py   front-ends; they use only the enabled features
  core/                   shared parts: settings, errors, refs, paging, send log, text, markup and time helpers
  backends/outlook/       COM session (one STA thread, time-out), constants, free/busy, profile (signatures, fonts)
  backends/teams/         DevTools connection to the Teams browser, page script (hook.js), Teams session
  backends/fake/          fake backend (JSON state), test data, snapshot of real data
  features/<name>/        api.py, service.py (Outlook), fake.py, models.py, rules.py, render.py, cli.py, mcp.py
```

## Rules

* A feature uses `core` and `backends`. A feature never uses another feature.
* `api.py` defines the interface. `service.py` (Outlook or Teams) and `fake.py` (fake backend) supply it.
* `rules.py` has the logic that all backends share. `render.py` makes the short text.
* `cli.py` and `mcp.py` are thin adapters. They do not contain logic.
* A backend starts at its first use. A command that does not use Outlook does not start Outlook.

## The Teams backend

The tool starts a separate browser profile with a DevTools port on 127.0.0.1 and opens Teams.
A page script (`hook.js`) keeps the tokens that the Teams web app gets for your home account.
`TeamsSession.request()` runs a `fetch()` inside the tab, so the tokens never leave the browser.
The page script only calls the Teams service addresses from the service map of your account.

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
