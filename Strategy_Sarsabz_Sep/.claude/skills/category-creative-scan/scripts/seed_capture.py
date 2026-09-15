"""Seed capture.xlsx with one row per collected post, ready to code.

    python seed_capture.py <workspace>/scan

Closes the gap between collection and coding: `collect_ig.py` writes posts.csv, and
`analyze_scan.py` reads the Ads sheet, but nothing carried the provenance across. Without
this you hand-transcribe 80+ rows of URLs and filenames, which is both slow and exactly
the sort of thing that loses a `source_url` and silently drops the row at analysis.

Fills only what the collector actually knows: post_id, brand, ring, handle, date, format,
reposted_from, source_url, image_file, and the caption/alt text into `notes` so you can
read it while coding. Every judgement column is left empty for you.

Re-runnable. Existing rows are matched on post_id and their codes preserved, so you can
collect more brands and re-seed without losing work.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))
from make_handles import COLS  # noqa: E402  single-sourced schema

# Filled from the collector; everything else is your judgement.
FROM_COLLECTOR = {"post_id", "brand", "ring", "handle", "date", "pinned_guess",
                  "format", "reposted_from", "source_url", "image_file", "notes"}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workspace", type=Path)
    ap.add_argument("--posts", type=Path, default=None,
                    help="posts.csv to seed from (default: <workspace>/posts.csv)")
    args = ap.parse_args()

    ws_dir = args.workspace
    posts_path = args.posts or ws_dir / "posts.csv"
    cap_path = ws_dir / "capture.xlsx"
    if not posts_path.exists():
        sys.exit(f"{posts_path} not found - run collect_ig.py first.")
    if not cap_path.exists():
        sys.exit(f"{cap_path} not found - run make_handles.py first.")

    with posts_path.open(encoding="utf-8", newline="") as fh:
        posts = list(csv.DictReader(fh))
    if not posts:
        sys.exit(f"{posts_path} has no rows.")

    wb = openpyxl.load_workbook(cap_path)
    ws = wb["Ads"]
    hdr = [c.value for c in ws[1]]
    idx = {name: i for i, name in enumerate(hdr)}

    # keep any coding already done, matched on post_id
    existing: dict[str, list] = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[idx["post_id"]]:
            existing[str(r[idx["post_id"]])] = list(r)

    ws.delete_rows(2, max(ws.max_row - 1, 1))

    # The collector emits posts in grid order. Instagram serves up to 3 PINNED posts
    # first, and a pin can be months older than the rest - which wrecks cadence. Prefer
    # the collector's own flag; fall back to position when seeding older output.
    seen_per_handle: dict[str, int] = {}

    written = kept = 0
    for p in posts:
        sc = p.get("shortcode") or p.get("post_id") or ""
        if not sc:
            continue
        handle = p.get("handle", "")
        pos = seen_per_handle.get(handle, 0)
        seen_per_handle[handle] = pos + 1
        pinned = (p.get("pinned_guess") or ("yes" if pos < 3 else "no"))

        row = existing.get(sc, [None] * len(hdr))
        note_bits = []
        if p.get("caption"):
            note_bits.append(f"CAPTION: {p['caption']}")
        elif p.get("alt_text"):
            note_bits.append(f"ALT: {p['alt_text']}")
        if pinned == "yes":
            note_bits.append("[possibly pinned - excluded from cadence span]")
        if p.get("reposted_from"):
            note_bits.append(f"[repost from @{p['reposted_from']}]")
        note = "  ".join(note_bits)[:1000]

        vals = {
            "post_id": sc, "brand": p.get("brand", ""), "ring": p.get("ring", ""),
            "handle": handle, "date": p.get("date", ""),
            "pinned_guess": pinned,
            "format": "video" if p.get("kind") == "reel" else "static",
            "reposted_from": p.get("reposted_from", ""),
            "source_url": p.get("source_url", ""),
            "image_file": p.get("image_file", ""),
            "notes": note,
        }
        out = []
        for i, name in enumerate(hdr):
            if name in FROM_COLLECTOR:
                out.append(vals.get(name, ""))
            else:
                out.append(row[i] if i < len(row) else None)
        ws.append(out)
        written += 1
        if sc in existing and any(row[idx[c]] for c in ("claim_primary", "content_type")
                                  if c in idx):
            kept += 1

    wb.save(cap_path)
    print(f"{cap_path}")
    print(f"  {written} row(s) seeded from {posts_path.name}")
    if kept:
        print(f"  {kept} row(s) had existing codes - preserved")
    blank = written - kept
    print(f"  {blank} row(s) awaiting coding. Read contact/ , then fill the judgement "
          f"columns.")


if __name__ == "__main__":
    main()
