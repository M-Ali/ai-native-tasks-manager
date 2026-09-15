"""Analyse the coded 24-month owned YouTube window -> owned_findings.md + CSVs.

Views are public YouTube view counts on 14-15 Sep 2026. On a brand channel most large counts
are bought (in-stream / Shorts ads), so views measure MEDIA WEIGHT the brand chose to put
behind an asset far more than organic interest. That is exactly why they are useful here:
they show where Sarsabz spends, not what farmers seek out. They are never share of voice.

`likely_paid` is an inference, labelled as one: >=100k views on a channel whose unpromoted
uploads typically draw hundreds to low thousands.
"""
import csv
import statistics
from collections import Counter, defaultdict

rows = list(csv.DictReader(open("codes_window.csv", encoding="utf-8")))
for r in rows:
    r["views"] = int(r["views"])
    r["length_sec"] = int(r["length_sec"])
    r["likely_paid"] = r["views"] >= 100_000

N = len(rows)
TV = sum(r["views"] for r in rows)
out = []
add = out.append


def pct(a, b):
    return f"{100 * a / b:.0f}%" if b else "-"


def table(key, title, csvname):
    cnt, vw = Counter(), Counter()
    for r in rows:
        cnt[r[key]] += 1
        vw[r[key]] += r["views"]
    add(f"\n### {title}\n")
    add(f"| {key} | items | share of items | views (M) | share of views |")
    add("|---|---:|---:|---:|---:|")
    data = []
    for k in sorted(cnt, key=lambda k: -vw[k]):
        add(f"| {k} | {cnt[k]} | {pct(cnt[k], N)} | {vw[k] / 1e6:.1f} | {pct(vw[k], TV)} |")
        data.append([k, cnt[k], vw[k]])
    with open(csvname, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow([key, "items", "views"])
        w.writerows(data)
    return cnt, vw


add("# Sarsabz owned YouTube - 24-month coded read")
add("")
dates = sorted(r["published"] for r in rows)
add(f"**Frame.** {N} uploads ({sum(r['tab'] == 'videos' for r in rows)} videos, "
    f"{sum(r['tab'] == 'shorts' for r in rows)} Shorts) on youtube.com/@Sarsabz published "
    f"{dates[0]} to {dates[-1]}; views read 14-15 Sep 2026; {TV / 1e6:.1f}M views in total. "
    "Coded from thumbnail, title and description - not watched end to end.")
add("")
add("**How to read views.** Brand-channel view counts are mostly bought. They show where the brand put "
    "media weight, not organic demand, and they are not share of voice. "
    f"{sum(r['likely_paid'] for r in rows)} of {N} uploads have >=100k views (inferred paid); "
    f"median of the rest is {statistics.median([r['views'] for r in rows if not r['likely_paid']]):,.0f}.")

table("platform", "By campaign platform", "by_platform.csv")
ct_cnt, ct_vw = table("content_type", "By content type", "by_content_type.csv")
cl_cnt, cl_vw = table("claim_primary", "By primary claim", "by_claim.csv")
table("proof_device", "By proof device", "by_proof.csv")
table("who_is_in_frame", "Who is in frame", "by_who.csv")
table("production_value", "Production value", "by_production.csv")
table("audience_generation", "Audience generation (inference from casting)", "by_generation.csv")

# functional vs everything else
func = [r for r in rows if r["claim_primary"] in ("yield_increase", "product_quality")]
fv = sum(r["views"] for r in func)
add("\n### The functional share\n")
add(f"Uploads whose primary claim is functional (yield or product quality): **{len(func)} of {N} "
    f"({pct(len(func), N)}), carrying {fv / 1e6:.1f}M of {TV / 1e6:.1f}M views ({pct(fv, TV)})**.")
proofed = [r for r in func if r["proof_device"] != "none"]
add(f"Of those, {len(proofed)} carry any proof device: "
    + ", ".join(f"#{r['n']} {r['title'][:40]} ({r['proof_device']})" for r in proofed) + ".")
yv = [r for r in rows if r["platform"] == "yield_10pct"]
add(f"The '10% se bhi ziyada' product films: {len(yv)} uploads, {sum(r['views'] for r in yv) / 1e6:.1f}M views; "
    f"{sum(r['proof_device'] == 'none' for r in yv)} of {len(yv)} offer no proof beyond the badge.")

# twin cuts: same message, long vs short
add("\n### Same message, two cuts\n")
add("| crop film | short cut views | long cut views | ratio |")
add("|---|---:|---:|---:|")
CROPS = {"gandum": "wheat", "makai": "maize", "gann": "sugarcane", "kapas": "cotton", "dhaan": "rice"}
by_crop = defaultdict(list)
for r in yv:
    crop = next((v for k, v in CROPS.items() if k in r["title"].lower()), None)
    if crop and "rameez" not in r["title"].lower():
        by_crop[crop].append(r)
for crop, rs in sorted(by_crop.items()):
    s = max((x for x in rs if x["length_sec"] <= 20), key=lambda x: x["views"], default=None)
    lg = max((x for x in rs if x["length_sec"] > 40), key=lambda x: x["views"], default=None)
    if s and lg:
        add(f"| {crop} | {s['views']:,} ({s['length_sec']}s) | {lg['views']:,} ({lg['length_sec']}s) | "
            f"{s['views'] / max(lg['views'], 1):,.0f}x |")
    else:
        add(f"| {crop} | " + " / ".join(f"{x['views']:,} ({x['length_sec']}s)" for x in rs) + " | single cut | - |")

# objection handling
obj = [r for r in rows if r["addresses_objection"] == "yes"]
add("\n### Addresses the switching objection\n")
add("Core objection, defined at scoping: *'why should I move from DAP + urea, which I know, to NP + CAN?'*")
add(f"Uploads that answer it with a real farmer's result: **{len(obj)} of {N}** - "
    + ", ".join(f"#{r['n']}" for r in obj) + ".")

# sponsorship vs product
sp = [r for r in rows if r["content_type"] == "sponsorship"]
add("\n### Sponsorship weight\n")
add(f"Multan Sultans content: {len(sp)} uploads ({pct(len(sp), N)}), {sum(r['views'] for r in sp) / 1e6:.1f}M views "
    f"({pct(sum(r['views'] for r in sp), TV)}). None shows a farmer or a product.")

# who is on camera for the brand's biggest assets
add("\n### Top 12 uploads by views\n")
add("| # | date | views | s | platform | claim | proof | who |")
add("|---:|---|---:|---:|---|---|---|---|")
for r in sorted(rows, key=lambda r: -r["views"])[:12]:
    add(f"| {r['n']} | {r['published']} | {r['views']:,} | {r['length_sec']} | {r['platform']} | "
        f"{r['claim_primary']} | {r['proof_device']} | {r['who_is_in_frame']} |")

open("owned_findings.md", "w", encoding="utf-8").write("\n".join(out) + "\n")
print("\n".join(out))
