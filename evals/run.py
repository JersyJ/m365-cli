from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
import tomllib
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[1]
DEFAULT_EXE = str(REPO / ".venv" / "Scripts" / "m365.exe")
DEFAULT_WORK = REPO / ".eval-runs"
CLAUDE_TIMEOUT_S = 900


@dataclass
class Metrics:
    """What one `claude -p` run cost and did."""

    tool_calls: list[str] = field(default_factory=list)
    tool_errors: int = 0
    denied: list[str] = field(default_factory=list)
    turns: int = 0
    input_tokens: int = 0
    output_tokens: int = 0
    cache_read_tokens: int = 0
    cost_usd: float = 0.0
    seconds: float = 0.0
    answer: str = ""


@dataclass
class Outcome:
    """Result of one task: the failed checks (empty = pass) and the metrics."""

    task: str
    failures: list[str]
    metrics: Metrics


def windows_path(path: Path) -> str:
    """Translate a WSL path for the Windows-side m365.exe."""
    return subprocess.run(
        ["wslpath", "-w", str(path)], check=True, capture_output=True, text=True
    ).stdout.strip()


def run_claude(prompt: str, mcp_config: Path, *, model: str | None, budget: float) -> Metrics:
    """Run one headless Claude Code session that can only use the m365 MCP server."""
    command = [
        "claude",
        "-p",
        prompt,
        "--mcp-config",
        str(mcp_config),
        "--strict-mcp-config",
        "--tools",
        "",
        "--allowedTools",
        "mcp__m365",
        "--output-format",
        "stream-json",
        "--verbose",
        "--no-session-persistence",
        "--max-budget-usd",
        str(budget),
    ]
    if model:
        command += ["--model", model]
    started = time.monotonic()
    proc = subprocess.run(command, capture_output=True, text=True, timeout=CLAUDE_TIMEOUT_S, check=False)
    metrics = Metrics(seconds=round(time.monotonic() - started, 1))
    for line in proc.stdout.splitlines():
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        _collect(event, metrics)
    if not metrics.answer and proc.returncode:
        metrics.answer = f"[claude exited {proc.returncode}] {proc.stderr.strip()[:500]}"
    return metrics


def _collect(event: dict[str, Any], metrics: Metrics) -> None:
    """Add the tool calls, errors, denials and final usage of one stream-json event."""
    if not isinstance(event, dict):
        return
    kind = event.get("type")
    message = event.get("message")
    content = message.get("content") if isinstance(message, dict) else None
    if not isinstance(content, list):
        content = []
    if kind == "assistant":
        metrics.tool_calls += [
            c["name"].removeprefix("mcp__m365__") for c in content if c.get("type") == "tool_use"
        ]
    elif kind == "user":
        metrics.tool_errors += sum(1 for c in content if c.get("type") == "tool_result" and c.get("is_error"))
    elif kind == "result":
        usage = event.get("usage") or {}
        metrics.answer = event.get("result") or ""
        metrics.turns = event.get("num_turns") or 0
        metrics.cost_usd = round(event.get("total_cost_usd") or 0.0, 4)
        metrics.input_tokens = usage.get("input_tokens", 0) + usage.get("cache_creation_input_tokens", 0)
        metrics.output_tokens = usage.get("output_tokens", 0)
        metrics.cache_read_tokens = usage.get("cache_read_input_tokens", 0)
        metrics.denied = [
            d.get("tool_name", "").removeprefix("mcp__m365__") for d in event.get("permission_denials") or []
        ]


def check(expect: dict[str, Any], state: dict[str, Any], original: dict[str, Any], m: Metrics) -> list[str]:
    """Compare the final fake state and the answer with the expectations of a task.

    Supported keys: answer_all (list of alternative groups), answer_none, outbox_empty, tools_used,
    denied_tools, draft, event_draft, flagged, changed_all, event_moved, canceled_removed, max_calls.
    See evals/README.md.
    """
    failures: list[str] = []
    answer = m.answer.lower()
    for group in expect.get("answer_all", []):
        if not any(alt.lower() in answer for alt in group):
            failures.append(f"answer lacks one of {group}")
    failures += [f"answer contains {bad!r}" for bad in expect.get("answer_none", []) if bad.lower() in answer]
    if expect.get("outbox_empty") and state["outbox"]:
        failures.append(f"something was sent: {[o['action'] for o in state['outbox']]}")
    failures += [f"tool {t} not used" for t in expect.get("tools_used", []) if t not in m.tool_calls]
    failures += [
        f"tool {t} was not stopped for approval" for t in expect.get("denied_tools", []) if t not in m.denied
    ]
    if spec := expect.get("draft"):
        failures += _check_draft(spec, state, original)
    if spec := expect.get("event_draft"):
        failures += _check_event_draft(spec, state, original)
    if spec := expect.get("flagged"):
        failures += _check_flagged(spec, state, original)
    if spec := expect.get("changed_all"):
        failures += _check_changed_all(spec, state, original)
    if spec := expect.get("event_moved"):
        failures += _check_event_moved(spec, state)
    if spec := expect.get("canceled_removed"):
        failures += _check_canceled_removed(spec, state, original)
    if (limit := expect.get("max_calls")) and len(m.tool_calls) > limit:
        failures.append(f"{len(m.tool_calls)} tool calls, expected at most {limit}")
    return failures


def _new_items(state: dict[str, Any], original: dict[str, Any], key: str) -> list[dict[str, Any]]:
    known = {item["id"] for item in original[key]}
    return [item for item in state[key] if item["id"] not in known]


def _check_draft(spec: dict[str, Any], state: dict[str, Any], original: dict[str, Any]) -> list[str]:
    drafts = [m for m in _new_items(state, original, "mails") if m["is_draft"]]
    subjects = {m["id"]: m["subject"].lower() for m in original["mails"]}
    for d in drafts:
        source = subjects.get(d["source_id"] or "", "")
        if spec.get("source_subject") and spec["source_subject"].lower() not in source:
            continue
        recipients = " ".join(d["to"] + d["cc"]).lower()
        if (
            spec.get("mode") in (None, d["draft_mode"])
            and spec.get("source_id") in (None, d["source_id"])
            and all(any(w.lower() in d["body"].lower() for w in group) for group in spec.get("body_all", []))
            and all(name.lower() in recipients for name in spec.get("to", []))
        ):
            return []
    return [f"no draft matching {spec}; drafts: {[(d['draft_mode'], d['subject'], d['to']) for d in drafts]}"]


def _overlaps(start: str, end: str, blocks: list[dict[str, Any]]) -> bool:
    return any(b["start"] < end and start < b["end"] and b["status"] in (1, 2, 3) for b in blocks)


def _check_event_draft(spec: dict[str, Any], state: dict[str, Any], original: dict[str, Any]) -> list[str]:
    events = _new_items(state, original, "events")
    matching = [e for e in events if spec.get("subject", "").lower() in e["subject"].lower()]
    if not matching:
        return [
            f"no new event with subject {spec.get('subject')!r}; new events: {[e['subject'] for e in events]}"
        ]
    e, failures = matching[0], []
    attendees = " ".join(e["required"] + e["optional"]).lower()
    failures += [f"attendee {n!r} missing" for n in spec.get("required", []) if n.lower() not in attendees]
    if spec.get("date") and not e["start"].startswith(spec["date"]):
        failures.append(f"event on {e['start'][:10]}, expected {spec['date']}")
    minutes = (datetime.fromisoformat(e["end"]) - datetime.fromisoformat(e["start"])).total_seconds() / 60
    if spec.get("minutes") and minutes != spec["minutes"]:
        failures.append(f"event is {minutes:.0f} min, expected {spec['minutes']}")
    if e["sent"]:
        failures.append("invitation was sent without approval")
    if spec.get("free"):
        mine = [
            {"start": x["start"], "end": x["end"], "status": x["show_as"]}
            for x in original["events"]
            if not x["canceled"]
        ]
        busy = [
            p["name"]
            for p in original["people"]
            if any(n.lower() in p["name"].lower() for n in spec.get("required", []))
            and _overlaps(e["start"], e["end"], p["blocks"] or [])
        ]
        if _overlaps(e["start"], e["end"], mine):
            busy.append("me")
        if busy:
            failures.append(f"slot {e['start'][11:16]}-{e['end'][11:16]} is busy for {busy}")
    return failures


def _check_flagged(spec: dict[str, Any], state: dict[str, Any], original: dict[str, Any]) -> list[str]:
    def wanted(m: dict[str, Any]) -> bool:
        return (
            spec.get("sender", "").lower() in m["sender_name"].lower()
            and spec.get("subject", "").lower() in m["subject"].lower()
            and (not spec.get("unread") or m["unread"])
        )

    before = {m["id"]: m for m in original["mails"]}
    targets = {i for i, m in before.items() if wanted(m)}
    flagged = {
        m["id"] for m in state["mails"] if m["flag"] == "set" and before.get(m["id"], {}).get("flag") != "set"
    }
    failures = []
    if missing := targets - flagged:
        failures.append(f"{len(missing)} target mail(s) not flagged: {sorted(missing)}")
    if extra := flagged - targets:
        failures.append(f"{len(extra)} other mail(s) flagged: {sorted(extra)}")
    return failures


def _check_changed_all(spec: dict[str, Any], state: dict[str, Any], original: dict[str, Any]) -> list[str]:
    """Every original mail matching sender/subject/folder now has `attr` == `value` (or contains it)."""
    targets = [
        m["id"]
        for m in original["mails"]
        if spec.get("sender", "").lower() in m["sender_name"].lower()
        and spec.get("subject", "").lower() in m["subject"].lower()
        and m["folder"] == spec.get("folder", m["folder"])
    ]
    if not targets:
        return [f"no mail matches {spec}"]
    now = {m["id"]: m for m in state["mails"]}
    attr, value = spec["attr"], spec["value"]
    wrong = [
        i
        for i in targets
        if not (value in now[i][attr] if isinstance(now[i][attr], list) else now[i][attr] == value)
    ]
    return [f"{len(wrong)}/{len(targets)} mail(s) without {attr}={value!r}: {wrong[:5]}"] if wrong else []


def _check_event_moved(spec: dict[str, Any], state: dict[str, Any]) -> list[str]:
    event = next((e for e in state["events"] if spec["subject"].lower() in e["subject"].lower()), None)
    if event is None:
        return [f"no event {spec['subject']!r}"]
    failures = []
    if not event["start"].startswith(spec["start"]):
        failures.append(f"event starts {event['start']}, expected {spec['start']}")
    if spec.get("minutes"):
        minutes = (
            datetime.fromisoformat(event["end"]) - datetime.fromisoformat(event["start"])
        ).total_seconds() / 60
        if minutes != spec["minutes"]:
            failures.append(f"event is {minutes:.0f} min, expected {spec['minutes']}")
    if spec.get("unsent") and event["sent"]:
        failures.append("update was sent without approval")
    return failures


def _check_canceled_removed(
    spec: dict[str, Any], state: dict[str, Any], original: dict[str, Any]
) -> list[str]:
    def in_range(e: dict[str, Any]) -> bool:
        return e["canceled"] and not e["organizer_is_me"] and spec["from"] <= e["start"][:10] <= spec["to"]

    before = [e["id"] for e in original["events"] if in_range(e)]
    left = [e["id"] for e in state["events"] if in_range(e)]
    others_lost = {e["id"] for e in original["events"] if not e["canceled"]} - {
        e["id"] for e in state["events"]
    }
    failures = []
    if not before:
        failures.append("no canceled meetings in range to start with")
    if left:
        failures.append(f"{len(left)}/{len(before)} canceled meeting(s) still there")
    if others_lost:
        failures.append(f"{len(others_lost)} active meeting(s) were removed too")
    return failures


def run_task(task: dict[str, Any], snapshot: Path, work: Path, args: argparse.Namespace) -> Outcome:
    """Prepare a fresh fake state for one task, run Claude on it, and check the result.

    Every task has its own folder, state file and MCP server process, so tasks can run in parallel.
    """
    folder = work / task["id"]
    folder.mkdir(parents=True, exist_ok=True)
    state_file = folder / "state.json"
    shutil.copyfile(snapshot, state_file)
    config = folder / "mcp.json"
    server = {
        "command": args.exe,
        "args": ["--fake-state", windows_path(state_file), "--data-dir", windows_path(folder), "mcp"],
    }
    config.write_text(json.dumps({"mcpServers": {"m365": server}}), encoding="utf-8")
    metrics = run_claude(task["prompt"], config, model=args.model, budget=args.budget)
    original = json.loads(snapshot.read_text(encoding="utf-8"))
    state = json.loads(state_file.read_text(encoding="utf-8"))
    return Outcome(task["id"], check(task.get("expect", {}), state, original, metrics), metrics)


def _line(outcome: Outcome) -> str:
    m = outcome.metrics
    status = "PASS" if not outcome.failures else "FAIL"
    head = (
        f"{status} {outcome.task:34} calls={len(m.tool_calls):2} errors={m.tool_errors} "
        f"denied={len(m.denied)} out_tok={m.output_tokens:5} ${m.cost_usd:.3f} {m.seconds:5.1f}s"
    )
    return "\n".join([head, *(f"     - {f}" for f in outcome.failures)])


def main() -> None:
    """Run evaluation suites in parallel, print one line per task, and write report.json."""
    parser = argparse.ArgumentParser(description="Evaluate the m365 MCP tools with headless Claude Code.")
    parser.add_argument(
        "--suite",
        nargs=2,
        action="append",
        metavar=("TASKS", "STATE"),
        default=[],
        help="Task file and fake state to start from; repeat for more suites.",
    )
    parser.add_argument("--tasks", type=Path, help="Shortcut for one suite (with --state).")
    parser.add_argument("--state", type=Path, help="Fake state for --tasks.")
    parser.add_argument("--only", default="", help="Comma-separated task ids.")
    parser.add_argument(
        "--model", default="haiku", help="Claude model (default haiku; use a larger one before releases)."
    )
    parser.add_argument("--jobs", type=int, default=6, help="Tasks running at the same time.")
    parser.add_argument("--budget", type=float, default=0.75, help="Max USD per task.")
    parser.add_argument("--exe", default=DEFAULT_EXE, help="Path of m365.exe.")
    parser.add_argument("--work", type=Path, default=DEFAULT_WORK, help="Directory for run outputs.")
    args = parser.parse_args()
    suites = [(Path(t), Path(st)) for t, st in args.suite]
    if args.tasks and args.state:
        suites.append((args.tasks, args.state))
    if not suites:
        parser.error("give --suite TASKS STATE (repeatable) or --tasks with --state")

    work = args.work / datetime.now().strftime("%Y%m%d-%H%M%S")
    jobs = []
    for tasks_file, state in suites:
        name = tasks_file.stem.removesuffix(".local")
        for task in tomllib.loads(tasks_file.read_text(encoding="utf-8"))["task"]:
            if not args.only or task["id"] in args.only.split(","):
                jobs.append(({**task, "id": f"{name}/{task['id']}"}, state))
    started = time.monotonic()
    outcomes: list[Outcome] = []
    with ThreadPoolExecutor(max_workers=max(1, args.jobs)) as pool:
        futures = [pool.submit(run_task, task, state, work, args) for task, state in jobs]
        for future in as_completed(futures):
            outcomes.append(future.result())
            print(_line(outcomes[-1]), flush=True)
    outcomes.sort(key=lambda o: o.task)
    passed = sum(not o.failures for o in outcomes)
    total_cost = sum(o.metrics.cost_usd for o in outcomes)
    elapsed = time.monotonic() - started
    print(
        f"\n{passed}/{len(outcomes)} passed, total ${total_cost:.2f}, {elapsed:.0f}s wall time "
        f"({args.model}, {args.jobs} jobs). Details: {work / 'report.json'}"
    )
    work.mkdir(parents=True, exist_ok=True)
    (work / "report.json").write_text(
        json.dumps([asdict(o) for o in outcomes], ensure_ascii=False, indent=1), encoding="utf-8"
    )
    sys.exit(0 if passed == len(outcomes) else 1)


if __name__ == "__main__":
    main()
