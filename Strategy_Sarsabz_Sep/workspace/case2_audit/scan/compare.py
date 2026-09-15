"""Sarsabz vs Engro on the same coded schema -> scan_findings.md + comparison CSVs.

Both windows are 24 months of owned YouTube, coded from thumbnail/title/description on the
category-creative-scan schema. Views are bought media weight, not organic interest, and the two
channels are not the same size - so shares within a brand are comparable, absolute views are not.
"""
import csv
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).parent
SETS = {
    "Sarsabz": HERE.parent / "owned" / "codes_window.csv",
    "Engro": HERE / "engro" / "codes_window.csv",
}
rows = {}
for brand, path in SETS.items():
    rs = list(csv.DictReader(path.open(encoding="utf-8")))
    for r in rs:
        r["views"] = int(r["views"])
        r["length_sec"] = int(r["length_sec"])
    rows[brand] = rs

out = []
add = out.append


def pct(a, b):
    return f"{100 * a / b:.0f}%" if b else "-"


def side_by_side(key, title):
    add(f"\n### {title}\n")
    keys = sorted({r[key] for rs in rows.values() for r in rs},
                  key=lambda k: -sum(1 for rs in rows.values() for r in rs if r[key] == k))
    add("| " + key + " | " + " | ".join(f"{b} items | {b} share" for b in rows) + " |")
    add("|---" + "|---:|---:" * len(rows) + "|")
    for k in keys:
        cells = []
        for b, rs in rows.items():
            n = sum(1 for r in rs if r[key] == k)
            cells += [str(n), pct(n, len(rs))]
        add(f"| {k} | " + " | ".join(cells) + " |")


add("# Sarsabz vs Engro — coded comparison of 24 months of owned YouTube")
add("")
for b, rs in rows.items():
    v = sum(r["views"] for r in rs)
    add(f"- **{b}**: {len(rs)} uploads, {v / 1e6:.1f}M views, "
        f"median length {sorted(r['length_sec'] for r in rs)[len(rs) // 2]}s, "
        f"longest {max(r['length_sec'] for r in rs)}s.")
add("")
add("Same schema, same window (15 Sep 2024 - 15 Sep 2026), same method: coded from thumbnail, title and "
    "description. Shares within a brand are comparable; absolute views are not, because media weight differs.")

side_by_side("content_type", "What kind of content each brand publishes")
side_by_side("claim_primary", "What each brand claims")
side_by_side("proof_device", "What each brand offers as proof")
side_by_side("who_is_in_frame", "Who is on screen")

# the two claims Engro makes and Sarsabz does not
add("\n### Claims only one brand makes\n")
for claim in ("cost_saving", "authenticity", "advisory_service", "yield_increase", "rewards_prizes",
              "national_identity", "farmer_recognition"):
    cells = []
    for b, rs in rows.items():
        n = sum(1 for r in rs if r["claim_primary"] == claim)
        vw = sum(r["views"] for r in rs if r["claim_primary"] == claim)
        cells.append(f"{n} ({vw / 1e6:.1f}M)")
    add(f"- **{claim}** — " + " · ".join(f"{b}: {c}" for b, c in zip(rows, cells)))

# expert / farmer authority
add("\n### Authority: who is allowed to speak\n")
for b, rs in rows.items():
    exp = sum(1 for r in rs if r["who_is_in_frame"] == "expert")
    cus = sum(1 for r in rs if r["who_is_in_frame"] == "customer")
    cel = sum(1 for r in rs if r["who_is_in_frame"] == "celebrity")
    named = sum(1 for r in rs if "named" in (r["notes"] or "").lower())
    add(f"- **{b}**: expert {exp}, customer {cus}, celebrity {cel}; "
        f"{named} uploads feature a *named* farmer or agronomist.")

# long-form
add("\n### Long-form\n")
for b, rs in rows.items():
    lf = [r for r in rs if r["length_sec"] >= 300]
    add(f"- **{b}**: {len(lf)} uploads of 5 minutes or longer"
        + (f", median {sorted(r['views'] for r in lf)[len(lf) // 2]:,} views" if lf else "")
        + (f"; longest {max(r['length_sec'] for r in lf)}s" if lf else "."))

# ecosystem
add("\n### Ecosystem push (app, retail, service)\n")
for b, rs in rows.items():
    eco = [r for r in rs if r["ecosystem_push"] == "yes"]
    add(f"- **{b}**: {len(eco)} of {len(rs)} uploads drive to an owned app, centre or service "
        f"({sum(r['views'] for r in eco) / 1e6:.1f}M views).")

# objection
add("\n### Answering the farmer's objection\n")
for b, rs in rows.items():
    ob = [r for r in rs if r["addresses_objection"] == "yes"]
    add(f"- **{b}**: {len(ob)} of {len(rs)} ({pct(len(ob), len(rs))}), {sum(r['views'] for r in ob) / 1e6:.1f}M views.")

with (HERE / "brand_compare.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["dimension", "value", *rows])
    for key in ("content_type", "claim_primary", "proof_device", "who_is_in_frame", "production_value"):
        for k in sorted({r[key] for rs in rows.values() for r in rs}):
            w.writerow([key, k, *[sum(1 for r in rs if r[key] == k) for rs in rows.values()]])

(HERE / "scan_findings.md").write_text("\n".join(out) + "\n", encoding="utf-8")
print("\n".join(out))
