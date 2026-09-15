"""Place every brand on the Brand Meaning Ladder from coded advertising - computed, not asserted.

    python place_brands.py <codes.csv|capture.xlsx> --map claim_to_level.yaml --out <workspace>/ladder

Reads one row per ad, maps its dominant claim to a level (1-6), and writes:

  brand_levels.csv   brand, n, mean_level, modal_level, and the count at each level
  level_mix.csv      the category's spread: how many ads sit at each level
  placement.md       the same, readable, with the caveats that belong on the slide

Why a script rather than a judgement call: the mean level per brand is the number the placement
slide rests on, and it gets quoted in a pitch room. It has to be reproducible on a rerun and
traceable to the rows that produced it.

An unmapped claim value is a FAILURE, not a skipped row. Dropping rows silently moves every mean
in the table, and the error is invisible in the output.
"""

from __future__ import annotations

import argparse
import csv
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

THIN_N = 8  # below this, a brand's mean is an anecdote and is flagged as such


def read_rows(path: Path) -> list[dict]:
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl

        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        rows = list(ws.iter_rows(values_only=True))
        head = [str(h).strip() if h is not None else "" for h in rows[0]]
        return [dict(zip(head, ["" if c is None else str(c) for c in r])) for r in rows[1:]]
    with path.open(encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("codes", type=Path, help="coded ads: codes.csv or capture.xlsx")
    ap.add_argument("--map", type=Path, required=True, help="claim_to_level.yaml")
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--client", default=None, help="client brand, marked in the output")
    args = ap.parse_args()

    import yaml

    cfg = yaml.safe_load(args.map.read_text(encoding="utf-8"))
    claim_col = cfg.get("column", "claim_primary")
    brand_col = cfg.get("brand_column", "brand")
    level_col = cfg.get("level_column", "level")
    levels = {int(k): v for k, v in (cfg.get("levels") or {}).items()}
    mapping = {str(k): int(v) for k, v in (cfg.get("map") or {}).items()}
    exclude = cfg.get("exclude") or {}

    rows = read_rows(args.codes)
    if not rows:
        sys.exit(f"{args.codes} has no rows.")
    for col in (claim_col, brand_col):
        if col not in rows[0]:
            sys.exit(f"{args.codes} has no '{col}' column. Columns: {', '.join(rows[0])}")

    kept, dropped = [], Counter()
    for r in rows:
        skip = False
        for col, values in exclude.items():
            if r.get(col, "") in values:
                dropped[f"{col}={r.get(col)}"] += 1
                skip = True
        if not skip:
            kept.append(r)

    unmapped = sorted({r[claim_col] for r in kept
                       if not r.get(level_col) and r[claim_col] not in mapping})
    if unmapped:
        sys.exit("Unmapped claim values (add them to the map, or exclude them explicitly):\n  "
                 + "\n  ".join(unmapped))

    by_brand: dict[str, list[int]] = defaultdict(list)
    overall = Counter()
    for r in kept:
        lvl = int(r[level_col]) if r.get(level_col) else mapping[r[claim_col]]
        by_brand[r[brand_col] or "(unnamed)"].append(lvl)
        overall[lvl] += 1

    args.out.mkdir(parents=True, exist_ok=True)
    order = sorted(by_brand, key=lambda b: statistics.fmean(by_brand[b]))

    with (args.out / "brand_levels.csv").open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["brand", "n", "mean_level", "modal_level", "thin_sample"]
                   + [f"L{i}" for i in range(1, 7)])
        for b in order:
            v = by_brand[b]
            c = Counter(v)
            w.writerow([b, len(v), f"{statistics.fmean(v):.1f}", Counter(v).most_common(1)[0][0],
                        "yes" if len(v) < THIN_N else ""] + [c.get(i, 0) for i in range(1, 7)])

    with (args.out / "level_mix.csv").open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["level", "name", "ads", "share_pct"])
        total = sum(overall.values())
        for i in range(1, 7):
            w.writerow([i, levels.get(i, ""), overall.get(i, 0),
                        f"{100 * overall.get(i, 0) / total:.1f}" if total else "0"])

    total = sum(overall.values())
    cat_mean = statistics.fmean([lvl for b in by_brand for lvl in by_brand[b]])
    out = [f"# Ladder placement — {args.codes.name}", "",
           f"**{total} coded ads across {len(by_brand)} brands.** Category mean level "
           f"**{cat_mean:.1f}**; modal level **{overall.most_common(1)[0][0]}**.", ""]
    if dropped:
        out.append("Excluded by rule: " + ", ".join(f"{k} ({n})" for k, n in dropped.items()) + ".")
        out.append("")
    out += ["| Brand | n | mean | mode | " + " | ".join(f"L{i}" for i in range(1, 7)) + " |",
            "|---|---:|---:|---:|" + "---:|" * 6]
    for b in order:
        v = by_brand[b]
        c = Counter(v)
        flag = " *(thin)*" if len(v) < THIN_N else ""
        star = " **(client)**" if args.client and b == args.client else ""
        out.append(f"| {b}{star}{flag} | {len(v)} | {statistics.fmean(v):.1f} | "
                   f"{c.most_common(1)[0][0]} | " + " | ".join(str(c.get(i, 0)) for i in range(1, 7)) + " |")
    out += ["", "## Levels", ""] + [f"- **{i}. {levels.get(i, '')}** — {overall.get(i, 0)} ads"
                                    for i in range(1, 7)]
    out += ["",
            "**Read this as a claim map, not a quality ranking.** Higher is not better: the target is the "
            "deepest level a brand can prove. A mean is only meaningful with its n — brands marked *(thin)* "
            f"have fewer than {THIN_N} coded ads. Placement reflects what was collected, which on a capped "
            "sample is a current-period read, not a history."]
    (args.out / "placement.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    print(f"{total} ads, {len(by_brand)} brands -> {args.out}")
    print(f"  category mean level {cat_mean:.1f}")
    for b in order:
        v = by_brand[b]
        print(f"  {b:28} n={len(v):4}  mean={statistics.fmean(v):.1f}"
              + ("  (thin)" if len(v) < THIN_N else ""))


if __name__ == "__main__":
    main()
