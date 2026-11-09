# Evaluations

The evaluations run real Claude Code sessions (`claude -p`) with the m365 MCP server.
The server uses the fake backend. Thus, the sessions never change your real mailbox.
After each session, `run.py` reads the fake state and checks the result.

## Data

There are two data sets:

* Synthetic data (`m365_cli/backends/fake/fixtures.py`). Tasks: `tasks/fixtures.toml`.
* A snapshot of your real Outlook data. Tasks: `tasks/real.local.toml`.

The snapshot and the `*.local.toml` files contain private data. Git ignores them. Do not share them.

## Run the evaluations

Run these steps in WSL, in the project directory:

1. Make the synthetic state:
   `m365 dev reset-fake 'C:\Users\<you>\AppData\Local\m365-cli\fake-fixtures.json'`
2. Make the snapshot of your real data (takes approximately 2 minutes):
   `m365 dev snapshot --with-injection-mail`
3. Run both suites in one run:
   `python3 evals/run.py --suite evals/tasks/fixtures.toml <fake-fixtures.json> --suite evals/tasks/real.local.toml <fake-snapshot.json>`

The tasks run in parallel (`--jobs`, default 6). The default model is Haiku.
A full run of 30 tasks takes approximately 1 minute and costs approximately 0.05 USD.

* Use `--model sonnet` or `--model opus` before a release. The larger models are slower and cost more.
* Use `--only real/thread-task,fixtures/move-meeting` to run some tasks.
* Each task has a cost limit. The default is 0.75 USD (`--budget`).

## Results

`run.py` shows one line for each task: PASS or FAIL, tool calls, tool errors, denied calls, output tokens, cost and time.
The file `report.json` in the run directory (`.eval-runs/<time>/`) contains all details and the answers.

Examine these values after each change to a tool or a description:

* Many tool calls: the tool needs better filters or paging.
* Tool errors: the description of a parameter is not clear.
* Denied calls: a send tool was stopped for approval. This is correct for send tasks.

## Task format

```toml
[[task]]
id = "reply-draft"
prompt = "Reply to Petr ... Only prepare it, do not send."
[task.expect]
draft = { mode = "reply", source_id = "m2", body_all = [["friday"]] }
outbox_empty = true
```

| Key            | Check |
|----------------|-------|
| `answer_all`   | The answer contains one word of each group. |
| `answer_none`  | The answer does not contain these words. |
| `outbox_empty` | Nothing was sent. |
| `tools_used`   | These tools were called. |
| `denied_tools` | These tools were stopped for approval. |
| `draft`        | A new draft exists: `mode`, `source_id` or `source_subject`, `to`, `body_all`. |
| `event_draft`  | A new unsent event exists: `subject`, `required`, `date`, `minutes`, `free` (no overlap with busy times). |
| `flagged`      | Exactly the mails that agree with `sender`, `subject` and `unread` are now flagged. |
| `changed_all`  | All mails that agree with `sender`, `subject` and `folder` now have `attr` = `value`. |
| `event_moved`  | The event with `subject` now starts at `start`, has `minutes`, and is not sent (`unsent`). |
| `canceled_removed` | No canceled meeting of other people remains between `from` and `to`. No active meeting is removed. |
| `max_calls`    | The session used this number of tool calls or fewer. |

Real data changes between snapshots. Thus, do not use item IDs in real tasks. Use `source_subject`.
