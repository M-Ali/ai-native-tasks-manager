"""Turn a messy YouTube-comment export into clean rows, and say what was discarded.

    python load_comments.py <file.xlsx|file.txt|file.csv> --out <dir>

Writes comments.csv (text, source_col, row, is_reply, likes, author, posted, lang_script)
and load_report.md, which records how many rows were read, how many were scaffolding,
and what was dropped. Read that report before quoting any number - if the parser
misread the layout, the report is where it shows.

WHY THIS EXISTS

Pasted YouTube exports are not tabular. Two layouts turn up, often in the same file:

  BLOCK              FLAT
  row: comment text  row: comment text
  row: 12            row: comment text
  row: Reply         row: comment text
  row: @handle
  row: 3 months ago
  row: Hide replies

In block layout the like count, author and timestamp are separate ROWS beneath the
comment, not columns beside it. Treating every non-empty cell as a comment inflates the
corpus by 5-6x and fills the analysis with rows reading "Reply", "1" and "@username".
One real file had 828 non-empty cells in column A and ~805 actual comments; another
could be 6,000 cells and 1,000 comments. Never report the raw row count as the sample.

A second column often holds REPLIES, sometimes with column A empty on that row. Replies
are real consumer voice and are kept, flagged `is_reply`, because a reply is usually a
rebuttal or a corroboration and that is exactly what you want when reading a debate.

Scaffolding is recognised by shape, never by position, because the two layouts
interleave. See SCAFFOLD_EXACT / the regexes below.
"""

from __future__ import annotations

import argparse
import csv
import html
import re
import sys
from collections import Counter
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# UI chrome that YouTube's own page emits; never consumer text.
SCAFFOLD_EXACT = {
    "reply", "replies", "hide replies", "show replies", "view replies",
    "comments", "comment", "top comments", "newest first", "sort by",
    "pinned by creator", "pinned", "read more", "show more", "show less",
    "cancel", "subscribe", "subscribed", "like", "dislike", "share",
    "author", "likes", "timestamp", "date", "text", "content", "message",
}
SCAFFOLD_PREFIX = ("top is selected", "sort by", "@")          # @handle = author row
RE_AGO = re.compile(r"^\s*\d+\s*(second|minute|hour|day|week|month|year)s?\s+ago\s*$", re.I)
RE_COUNT = re.compile(r"^\s*[\d,.]+\s*[KM]?\s*$", re.I)        # bare like count
RE_REPLYCOUNT = re.compile(r"^\s*\d+\s+repl(y|ies)\s*$", re.I)
RE_URDU = re.compile(r"[؀-ۿ]")
RE_LATIN = re.compile(r"[A-Za-z]")


def clean(raw: object) -> str:
    """Normalise a cell to plain text. Exports carry literal <br> and HTML entities."""
    s = str(raw)
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]{1,40}>", " ", s)
    s = html.unescape(s)
    s = s.replace("​", "").replace("﻿", "")
    return re.sub(r"\s+", " ", s).strip()


def is_scaffold(t: str) -> bool:
    low = t.lower().strip()
    if low in SCAFFOLD_EXACT or not low:
        return True
    if low.startswith(SCAFFOLD_PREFIX):
        return True
    return bool(RE_AGO.match(t) or RE_COUNT.match(t) or RE_REPLYCOUNT.match(t))


def script_of(t: str) -> str:
    ur, la = bool(RE_URDU.search(t)), bool(RE_LATIN.search(t))
    if ur and la:
        return "mixed"
    if ur:
        return "urdu_script"
    if la:
        return "latin"
    return "other"


def read_grid(path: Path) -> list[list[str]]:
    """Every input becomes a grid of strings so one parser handles all formats."""
    suf = path.suffix.lower()
    if suf in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        grid = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                grid.append(["" if c is None else clean(c) for c in row])
        return grid
    if suf == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as fh:
            return [[clean(c) for c in r] for r in csv.reader(fh)]
    # .txt and anything else: one comment per line, blank lines separate
    text = path.read_text(encoding="utf-8", errors="replace")
    return [[clean(l)] for l in text.splitlines()]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--min-chars", type=int, default=2,
                    help="drop shorter than this; emoji-only reactions carry no codeable "
                         "subject but ARE counted in the report")
    args = ap.parse_args()

    if not args.source.exists():
        sys.exit(f"{args.source} not found")
    args.out.mkdir(parents=True, exist_ok=True)
    grid = read_grid(args.source)

    rows, dropped, short = [], 0, 0
    header_skipped = False
    for i, r in enumerate(grid):
        for col, cell in enumerate(r):
            if not cell:
                continue
            # a first-row header like "COMMENTS | REPLY" is a label, not a comment
            if i == 0 and cell.lower() in SCAFFOLD_EXACT:
                header_skipped = True
                continue
            if is_scaffold(cell):
                dropped += 1
                continue
            if len(cell) < args.min_chars:
                short += 1
                continue
            rows.append({
                "text": cell,
                "source_col": col,
                "row": i + 1,
                # Column 0 is the main thread; any further column is a reply pane.
                "is_reply": "yes" if col > 0 else "no",
                "lang_script": script_of(cell),
            })

    # Block layout only: the like count and author sit directly under the comment.
    # Attach them where the shape is unambiguous, leave blank otherwise - a guessed
    # like count is worse than an empty column.
    flat = {(x["row"], x["source_col"]): x for x in rows}
    for x in rows:
        x["likes"] = x.get("likes", "")
        x["author"] = x.get("author", "")
        x["posted"] = x.get("posted", "")
        if x["source_col"] != 0:
            continue
        nxt = [grid[j][0] if j < len(grid) and grid[j] else "" for j in
               range(x["row"], min(x["row"] + 4, len(grid)))]
        if nxt and RE_COUNT.match(nxt[0] or "") and (x["row"] + 1, 0) not in flat:
            x["likes"] = nxt[0].strip()
        for cand in nxt:
            if cand.startswith("@") and not x["author"]:
                x["author"] = cand.strip()
            if RE_AGO.match(cand or "") and not x["posted"]:
                x["posted"] = cand.strip()

    # Exports repeat a comment when it appears in several reply panes. Seven copies of
    # one sentence is ONE person, and reporting it as seven would invent six consumers.
    # Flag it rather than dropping it: a phrase genuinely typed by many people is a real
    # signal, and only a human reading the rows can tell the two cases apart.
    text_counts = Counter(x["text"] for x in rows)
    for x in rows:
        x["times_in_file"] = text_counts[x["text"]]
    dup_rows = sum(1 for x in rows if x["times_in_file"] > 1)
    dup_texts = sum(1 for v in text_counts.values() if v > 1)

    out_csv = args.out / "comments.csv"
    cols = ["text", "source_col", "row", "is_reply", "likes", "author", "posted",
            "lang_script", "times_in_file"]
    with out_csv.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for x in rows:
            w.writerow({k: x.get(k, "") for k in cols})

    total_cells = sum(1 for r in grid for c in r if c)
    replies = sum(1 for x in rows if x["is_reply"] == "yes")
    scripts: dict[str, int] = {}
    for x in rows:
        scripts[x["lang_script"]] = scripts.get(x["lang_script"], 0) + 1

    rep = args.out / "load_report.md"
    lines = [
        "# Comment load report", "",
        f"- Source: `{args.source.name}`",
        f"- Rows in file: **{len(grid)}**",
        f"- Non-empty cells: **{total_cells}**",
        f"- Interface scaffolding discarded (Reply / Hide replies / @handle / "
        f"timestamps / bare counts): **{dropped}**",
        f"- Too short to code (< {args.min_chars} chars, mostly emoji-only): **{short}**",
        f"- **Comments analysed: {len(rows)}** ({replies} of them replies)",
        f"- Rows whose text repeats elsewhere in the file: **{dup_rows}** "
        f"across {dup_texts} distinct text(s)",
        "",
        "## Script mix", "",
        "| Script | Comments |", "|---|---:|",
    ]
    for k, v in sorted(scripts.items(), key=lambda kv: -kv[1]):
        lines.append(f"| {k} | {v} |")
    lines += [
        "",
        "Quote **comments analysed**, never the raw row or cell count - the difference "
        "between them is interface scaffolding, not consumer voice.",
        "",
        "## Repeated text", "",
        "`times_in_file` counts how often each exact text appears. Exports repeat a "
        "comment when it shows in several reply panes, so **N copies is usually one "
        "person, not N people** - check the rows before treating a repeat as "
        "corroboration. A short phrase many people genuinely type is a real signal; a "
        "long identical sentence is almost always the same comment captured twice.",
    ]
    if header_skipped:
        lines.append("")
        lines.append("A header row was detected and excluded.")
    rep.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"{out_csv}: {len(rows)} comment(s) from {total_cells} non-empty cell(s)")
    print(f"  {dropped} scaffolding, {short} too short, {replies} replies")
    print(f"{rep}: read this before quoting any figure")


if __name__ == "__main__":
    main()
