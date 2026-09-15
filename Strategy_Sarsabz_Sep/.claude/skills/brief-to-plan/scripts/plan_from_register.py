"""Turn requirements.yaml into a dated plan: what starts today, and what is already late.

Usage:
    python plan_from_register.py requirements.yaml --out <dir> [--today YYYY-MM-DD]
                                 [--buffer-days 1]

Back-plans from the submission deadline through the dependency graph. Each item's
duration is effort_days + lead_time_days, because time spent waiting on a bank or a
regulator is time you cannot compress by working harder.

Writes plan.md and register.xlsx.
"""

from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FILL = {
    "overdue": PatternFill("solid", fgColor="F8CBAD"),
    "today": PatternFill("solid", fgColor="FFE699"),
    "later": PatternFill("solid", fgColor="C6E0B4"),
}


def parse_deadline(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise SystemExit(f"Could not read submission_deadline: {value!r} (use YYYY-MM-DD HH:MM)")


def topo_order(reqs: dict[str, dict]) -> list[str]:
    """Dependencies first. Raises on a cycle rather than looping forever."""
    order: list[str] = []
    state: dict[str, int] = {}

    def visit(node: str, trail: list[str]) -> None:
        if state.get(node) == 2:
            return
        if state.get(node) == 1:
            raise SystemExit(f"Circular dependency: {' -> '.join([*trail, node])}")
        state[node] = 1
        for dep in reqs[node].get("depends_on") or []:
            if dep not in reqs:
                raise SystemExit(f"{node}: depends_on unknown id {dep!r}")
            visit(dep, [*trail, node])
        state[node] = 2
        order.append(node)

    for key in reqs:
        visit(key, [])
    return order


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("register", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--today", default=None, help="override today, for what-if planning")
    ap.add_argument("--buffer-days", type=int, default=1,
                    help="finish this many days before the deadline (default 1)")
    args = ap.parse_args()

    data = yaml.safe_load(args.register.read_text(encoding="utf-8")) or {}
    brief = data.get("brief") or {}
    items = data.get("requirements") or []
    if not items:
        raise SystemExit("No requirements in the register.")

    today = date.fromisoformat(args.today) if args.today else date.today()
    deadline = parse_deadline(brief.get("submission_deadline"))
    ready_by = deadline - timedelta(days=args.buffer_days)
    days_left = (deadline - today).days

    reqs = {}
    for item in items:
        key = item.get("id")
        if not key:
            raise SystemExit(f"Requirement without an id: {item.get('name')!r}")
        if key in reqs:
            raise SystemExit(f"Duplicate requirement id: {key!r}")
        reqs[key] = item

    order = topo_order(reqs)
    successors: dict[str, list[str]] = {k: [] for k in reqs}
    for key, item in reqs.items():
        for dep in item.get("depends_on") or []:
            successors[dep].append(key)

    # Backward pass: latest finish is the earliest latest-start among successors.
    latest_start: dict[str, date] = {}
    latest_finish: dict[str, date] = {}
    for key in reversed(order):
        item = reqs[key]
        duration = int(item.get("effort_days", 1)) + int(item.get("lead_time_days", 0))
        succ = successors[key]
        finish = min((latest_start[s] for s in succ), default=ready_by)
        # An item can have a deadline of its own that is earlier than the submission --
        # a clarification window, a site visit, a registration cut-off.
        if item.get("due"):
            finish = min(finish, parse_deadline(item["due"]))
        latest_finish[key] = finish
        latest_start[key] = latest_finish[key] - timedelta(days=max(duration, 0))

    for key in reqs:
        reqs[key]["_start"] = latest_start[key]
        reqs[key]["_finish"] = latest_finish[key]
        reqs[key]["_slack"] = (latest_start[key] - today).days
        reqs[key]["_duration"] = (
            int(reqs[key].get("effort_days", 1)) + int(reqs[key].get("lead_time_days", 0))
        )

    ranked = sorted(reqs.values(), key=lambda r: (r["_slack"], -int(r.get("marks", 0) or 0)))
    overdue = [r for r in ranked if r["_slack"] < 0]
    start_now = [r for r in ranked if r["_slack"] == 0]
    gates = [r for r in ranked if r.get("type") == "gate"]
    external = [r for r in ranked if r.get("external")]
    min_slack = min((r["_slack"] for r in reqs.values()), default=0)
    critical = [r for r in ranked if r["_slack"] == min_slack]

    # --- plan.md ----------------------------------------------------------
    L: list[str] = []
    add = L.append
    title = brief.get("title") or args.register.stem
    add(f"# {title} - plan")
    add("")
    if brief.get("client"):
        add(f"**Client:** {brief['client']}  ")
    if brief.get("reference"):
        add(f"**Reference:** {brief['reference']}  ")
    add(f"**Deadline:** {brief.get('submission_deadline')}  ")
    if brief.get("submission_channel"):
        add(f"**Channel:** {brief['submission_channel']}  ")
    if brief.get("clarification_deadline"):
        add(f"**Clarifications close:** {brief['clarification_deadline']}  ")
    add(f"**Planned from:** {today.isoformat()} - **{days_left} day(s) to deadline**, "
        f"working back from {ready_by.isoformat()} ({args.buffer_days}-day buffer)")
    add("")

    if data.get("the_ask"):
        add("## The ask")
        add("")
        add(str(data["the_ask"]).strip())
        add("")

    missing = data.get("missing") or []
    if missing:
        add("## Referenced but missing - resolve before anything else")
        add("")
        for m in missing:
            add(f"- **{m.get('what')}** - referenced at {m.get('referenced_at', 'n/a')}.  ")
            add(f"  Action: {m.get('action', 'confirm with the buyer')}")
            if m.get("question"):
                add(f"  Question to file: *{m['question']}*")
        if brief.get("clarification_deadline"):
            add("")
            add(f"Clarification window closes **{brief['clarification_deadline']}**. After that "
                "these become assumptions you carry, not questions you can ask.")
        add("")

    if overdue:
        add("## Already late")
        add("")
        add("These needed to start before today. Compress, escalate, or drop them deliberately -")
        add("do not simply carry them forward.")
        add("")
        add("| Requirement | Type | Marks | Days late | Owner |")
        add("|---|---|---:|---:|---|")
        for r in overdue:
            add(f"| {r.get('name')} | {r.get('type')} | {r.get('marks', 0)} | "
                f"{-r['_slack']} | {r.get('owner', '-')} |")
        add("")

    add("## Start today")
    add("")
    if start_now or overdue:
        add("| Requirement | Type | Marks | Effort | Waiting on others | Owner |")
        add("|---|---|---:|---:|---:|---|")
        for r in overdue + start_now:
            add(f"| {r.get('name')} | {r.get('type')} | {r.get('marks', 0)} | "
                f"{r.get('effort_days', 1)}d | {r.get('lead_time_days', 0)}d | "
                f"{r.get('owner', '-')} |")
    else:
        nxt = min(reqs.values(), key=lambda r: r["_start"])
        add(f"Nothing is due to start yet. Next up: **{nxt.get('name')}** on "
            f"{nxt['_start'].isoformat()}.")
    add("")

    if external:
        add("### Of those, the ones you do not control")
        add("")
        add("Third-party items set the schedule - you cannot finish them by working harder, so")
        add("they go first regardless of how few marks they carry.")
        add("")
        for r in external:
            flag = " **(late)**" if r["_slack"] < 0 else ""
            add(f"- {r.get('name')} - request by {r['_start'].isoformat()}, "
                f"{r.get('lead_time_days', 0)}d wait, {r.get('owner', 'unowned')}{flag}")
        add("")

    if gates:
        add("## Pass/fail gates")
        add("")
        add("Failing any one of these voids the submission whatever it scores elsewhere.")
        add("")
        add("| Gate | Start by | Marks | External | Owner |")
        add("|---|---|---:|---|---|")
        for r in sorted(gates, key=lambda r: r["_start"]):
            add(f"| {r.get('name')} | {r['_start'].isoformat()} | {r.get('marks', 0)} | "
                f"{'yes' if r.get('external') else 'no'} | {r.get('owner', '-')} |")
        add("")

    add("## Critical path")
    add("")
    add(f"Least slack in the plan is **{min_slack} day(s)**. Any slip on these moves the")
    add("submission date:")
    add("")
    for r in critical:
        add(f"- {r.get('name')} ({r['_duration']}d)")
    add("")

    add("## Full schedule")
    add("")
    add("| Start by | Finish by | Requirement | Type | Marks | Slack | Depends on | Owner |")
    add("|---|---|---|---|---:|---:|---|---|")
    for r in sorted(reqs.values(), key=lambda r: (r["_start"], -int(r.get("marks", 0) or 0))):
        deps = ", ".join(r.get("depends_on") or []) or "-"
        add(f"| {r['_start'].isoformat()} | {r['_finish'].isoformat()} | {r.get('name')} | "
            f"{r.get('type')} | {r.get('marks', 0)} | {r['_slack']} | {deps} | "
            f"{r.get('owner', '-')} |")
    add("")

    scored = [r for r in reqs.values() if r.get("type") == "scored" and int(r.get("marks", 0) or 0)]
    if scored:
        add("## Marks per day of effort")
        add("")
        add("Where effort buys the most score. Use it to allocate people **after** the gates and")
        add("external items are safe - never instead of them.")
        add("")
        add("| Requirement | Marks | Effort | Marks/day |")
        add("|---|---:|---:|---:|")
        for r in sorted(
            scored,
            key=lambda r: int(r.get("marks", 0)) / max(int(r.get("effort_days", 1)), 1),
            reverse=True,
        ):
            eff = max(int(r.get("effort_days", 1)), 1)
            add(f"| {r.get('name')} | {r.get('marks')} | {eff}d | "
                f"{int(r.get('marks', 0)) / eff:.1f} |")
        add("")
        total = sum(int(r.get("marks", 0) or 0) for r in reqs.values())
        add(f"Total marks modelled: **{total}**"
            + (f" of {brief['total_marks']} stated" if brief.get("total_marks") else "")
            + (f"; pass mark {brief['pass_mark']}." if brief.get("pass_mark") else "."))
        add("")

    add("## The call")
    add("")
    add("- **Bid / no-bid:** _decide and write it here_")
    add("- **Owners confirmed:** _named people, not functions_")
    add("- **Questions filed with the buyer:** _by when_")
    add("- **Deliberately not doing:** _what got dropped to fit the time_")
    add("")

    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "plan.md").write_text("\n".join(L), encoding="utf-8")

    # --- register.xlsx ----------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Register"
    headers = ["id", "name", "type", "marks", "effort_days", "lead_time_days", "external",
               "start_by", "finish_by", "slack_days", "depends_on", "owner", "notes"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for r, item in enumerate(
        sorted(reqs.values(), key=lambda x: (x["_start"], -int(x.get("marks", 0) or 0))), 2
    ):
        row = [
            item.get("id"), item.get("name"), item.get("type"), item.get("marks", 0),
            item.get("effort_days", 1), item.get("lead_time_days", 0),
            "yes" if item.get("external") else "no",
            item["_start"].isoformat(), item["_finish"].isoformat(), item["_slack"],
            ", ".join(item.get("depends_on") or []), item.get("owner", ""), item.get("notes", ""),
        ]
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            if c in (2, 11, 13):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        bucket = "overdue" if item["_slack"] < 0 else ("today" if item["_slack"] == 0 else "later")
        ws.cell(r, 10).fill = FILL[bucket]
        if item.get("type") == "gate":
            ws.cell(r, 3).font = Font(bold=True)
    for col, width in {1: 22, 2: 52, 3: 14, 4: 8, 5: 12, 6: 15, 7: 10, 8: 12, 9: 12,
                       10: 11, 11: 26, 12: 16, 13: 60}.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(reqs) + 1}"
    wb.save(args.out / "register.xlsx")

    print(f"Wrote {args.out / 'plan.md'} and {args.out / 'register.xlsx'}")
    print(f"{days_left} day(s) to deadline. {len(overdue)} overdue, {len(start_now)} start today, "
          f"{len(gates)} gate(s), critical-path slack {min_slack}d.")


if __name__ == "__main__":
    main()
