"""Affinity indices from a wide "Interests - free form" analytics export.

    python interest_index.py <interests.csv> --gender male --age 25-34,35-44,45-54 --out <dir>

The export has two header rows (gender, then age) above a row of segment totals, then one row per
interest. This computes, for each interest:

    index = (segment's share of that interest) / (whole file's share of that interest) x 100

WHY THE INDEX ALONE IS NOT THE ANSWER

An index says "this segment over-indexes here", not "this matters". An interest held by 40 people can
show i300 and mean nothing, while a 3,000-person interest at i160 is what a media planner can buy
against. So the output carries users AND index, small bases are excluded from the ranking rather than
silently ranked, and the dataset's own header (site, date range) is printed for the card's source line -
because the indices describe the population that produced the file, not the audience you have in mind.
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def load(path: Path):
    header_notes, rows = [], []
    for row in csv.reader(path.open(encoding="utf-8-sig")):
        if row and row[0].startswith("#"):
            note = row[0].lstrip("# ").strip()
            if note:
                header_notes.append(note)
        elif any(c.strip() for c in row):
            rows.append(row)
    if len(rows) < 4:
        sys.exit("This does not look like an interests export: fewer than 4 non-empty rows.")
    gender, age, _measure, base = rows[0], rows[1], rows[2], rows[3]
    cols = [i for i in range(1, len(gender) - 1) if (base[i] or "").strip().isdigit()]
    interests = []
    for r in rows[4:]:
        if not r or not r[0].strip():
            continue
        try:
            interests.append((r[0], [int((r[i] or 0) or 0) for i in cols], int(r[len(gender) - 1] or 0)))
        except (ValueError, IndexError):
            continue
    return header_notes, gender, age, base, cols, interests


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("csv_path", type=Path)
    ap.add_argument("--gender", default="", help="comma-separated, e.g. male (blank = all)")
    ap.add_argument("--age", default="", help="comma-separated, e.g. 25-34,35-44 (blank = all)")
    ap.add_argument("--min-users", type=int, default=50, help="below this, an index is noise (default 50)")
    ap.add_argument("--top", type=int, default=15)
    ap.add_argument("--out", type=Path, default=Path("."))
    args = ap.parse_args()

    notes, gender, age, base, cols, interests = load(args.csv_path)
    genders = [g.strip().lower() for g in args.gender.split(",") if g.strip()]
    ages = [a.strip() for a in args.age.split(",") if a.strip()]
    seg = [i for i in cols
           if (not genders or gender[i].strip().lower() in genders)
           and (not ages or age[i].strip() in ages)]
    if not seg:
        sys.exit(f"No column matches gender={args.gender!r} age={args.age!r}. "
                 f"Available: {sorted({(gender[i], age[i]) for i in cols})}")

    total_all = sum(int(base[i]) for i in cols)
    total_seg = sum(int(base[i]) for i in seg)
    idx_of = {i: n for n, i in enumerate(cols)}
    out = []
    for name, vals, grand in interests:
        seg_users = sum(vals[idx_of[i]] for i in seg)
        all_users = sum(vals[idx_of[i]] for i in cols) or grand
        if not all_users:
            continue
        index = round((seg_users / total_seg) / (all_users / total_all) * 100)
        out.append({"interest": name, "segment_users": seg_users, "all_users": all_users,
                    "segment_share_pct": round(100 * seg_users / total_seg, 2), "index": index,
                    "ranked": "yes" if seg_users >= args.min_users else "no (base too small)"})

    args.out.mkdir(parents=True, exist_ok=True)
    dest = args.out / "interest_index.csv"
    with dest.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(sorted(out, key=lambda r: -r["index"]))

    label = f"{args.gender or 'all genders'} {args.age or 'all ages'}"
    print(f"Segment: {label} - {total_seg:,} users of {total_all:,} ({100 * total_seg / total_all:.0f}% of the file)")
    if notes:
        print("Dataset header (put this on the card): " + " | ".join(notes))
        print("  These indices describe the population that produced this file. If that population is not your "
              "audience, say so on the card or leave the block out.")
    ranked = [r for r in out if r["ranked"] == "yes"]
    print(f"\nTop {args.top} by INDEX (min {args.min_users} users):")
    for r in sorted(ranked, key=lambda r: -r["index"])[:args.top]:
        print(f"  i{r['index']:<5} {r['segment_users']:>7,}  {r['interest']}")
    print(f"\nTop {args.top} by VOLUME - what a media plan can actually buy:")
    for r in sorted(ranked, key=lambda r: -r["segment_users"])[:args.top]:
        print(f"  i{r['index']:<5} {r['segment_users']:>7,}  {r['interest']}")
    small = len(out) - len(ranked)
    if small:
        print(f"\n{small} interests fell below {args.min_users} users and are unranked - their indices swing wildly.")
    print(f"\nwrote {dest}")


if __name__ == "__main__":
    main()
