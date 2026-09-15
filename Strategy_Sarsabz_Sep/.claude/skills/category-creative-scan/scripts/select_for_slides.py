"""Choose which posts appear on a brand's slide, by a stated rule.

    python select_for_slides.py <workspace>/scan --per-brand 12

Once you read more posts than a slide can hold, "which ones are shown" becomes a real
methodological question. Answering it with taste is the weakest link in an otherwise
evidenced section: an evaluator who asks "why these twelve?" should get a rule, not a
shrug.

The rule, in priority order:

1. **Posts carrying a proof device**, taken in rotation across distinct creative
   signatures rather than all at once. Proof devices are the most important ads in most
   categories - but where they are common, taking every one of them fills the slide with
   a single repeated template. `--max-per-signature` caps the lookalikes.
2. **Every post carrying a recurring campaign platform.** A brand with a platform is
   harder to displace than its follower count suggests, and the platform must be visible.
3. **The remainder, stratified to mirror that brand's coded content_type mix** across the
   full read - so the sheet represents what the brand actually publishes, not just what it
   happened to post most recently.

Within each stratum, most recent first. Writes `slide_selection.csv` and prints the mix so
the rule can be quoted on the method slide.
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workspace", type=Path)
    ap.add_argument("--per-brand", type=int, default=12)
    ap.add_argument("--max-per-signature", type=int, default=3,
                    help="cap on how many posts sharing one creative signature "
                         "(content_type + proof_device + claim) may be taken from the "
                         "proof-device tier, so a repeated template cannot fill a slide")
    ap.add_argument("--capture", type=Path, default=None)
    args = ap.parse_args()

    cap = args.capture or args.workspace / "capture.xlsx"
    if not cap.exists():
        sys.exit(f"{cap} not found.")

    wb = openpyxl.load_workbook(cap, data_only=True)
    ws = wb["Ads"]
    hdr = [c.value for c in ws[1]]
    rows = [dict(zip(hdr, [c.value for c in r])) for r in ws.iter_rows(min_row=2)
            if r[1].value]
    if not rows:
        sys.exit("No coded rows.")

    by_brand: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        by_brand[r["brand"]].append(r)

    out = []
    N = args.per_brand
    for brand, rs in by_brand.items():
        rs = sorted(rs, key=lambda r: str(r.get("date") or ""), reverse=True)
        chosen: list[dict] = []
        seen: set = set()

        def take(row, why):
            if row["post_id"] in seen or len(chosen) >= N:
                return
            seen.add(row["post_id"])
            chosen.append({**row, "_why": why})

        # Proof-bearing posts first - but ROUND-ROBIN across distinct creative
        # signatures rather than taking them all in date order.
        #
        # Taking them all only works where proof devices are rare. In categories
        # where they are common (a telco coding network_scale onto every coverage
        # tile), rule 1 swallows the whole slide and fills it with one repeated
        # template: a real scan put nine near-identical "NOW FIBERIZED <area>"
        # tiles on a brand's slide and dropped the distinctive series entirely,
        # because that series happened to carry no proof device.
        #
        # Rotating by signature keeps every proof TYPE visible while capping how
        # many lookalikes reach the page.
        def signature(r):
            return ((r.get("content_type") or "unclear"),
                    (r.get("proof_device") or "none"),
                    (r.get("claim_primary") or "unclear"))

        def round_robin(pool, why, cap=None):
            buckets: dict[tuple, list[dict]] = defaultdict(list)
            for r in pool:
                buckets[signature(r)].append(r)          # pool is already date-sorted
            order = sorted(buckets, key=lambda s: (s[1] == "none", -len(buckets[s])))
            depth = 0
            while len(chosen) < N and any(len(buckets[s]) > depth for s in order):
                if cap is not None and depth >= cap:
                    break
                for s in order:
                    if len(buckets[s]) > depth:
                        take(buckets[s][depth], why)
                depth += 1

        round_robin([r for r in rs
                     if r.get("proof_device") and r["proof_device"] != "none"],
                    "proof_device", cap=args.max_per_signature)
        for r in rs:
            if (r.get("campaign_platform") or "").strip():
                take(r, "campaign_platform")

        # stratify the remainder by the brand's real content-type mix
        remaining = [r for r in rs if r["post_id"] not in seen]
        mix = Counter(r.get("content_type") or "unclear" for r in rs)
        total = sum(mix.values()) or 1
        slots = N - len(chosen)
        if slots > 0 and remaining:
            quota = {k: max(1, round(slots * v / total)) for k, v in mix.items()}
            for ctype, q in sorted(quota.items(), key=lambda kv: -kv[1]):
                pool = [r for r in remaining if (r.get("content_type") or "unclear") == ctype]
                for r in pool[:q]:
                    take(r, f"stratified:{ctype}")
            for r in remaining:            # top up if rounding left slots
                take(r, "stratified:topup")

        for r in chosen:
            out.append({"brand": brand, "post_id": r["post_id"],
                        "image_file": r.get("image_file") or "",
                        "date": r.get("date") or "",
                        "content_type": r.get("content_type") or "",
                        "reason": r["_why"]})
        why = Counter(r["_why"].split(":")[0] for r in chosen)
        print(f"  {brand[:34]:36} {len(chosen):2}/{len(rs):3} selected  "
              f"({', '.join(f'{k} {v}' for k, v in why.most_common())})")

    path = args.workspace / "slide_selection.csv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["brand", "post_id", "image_file", "date",
                                           "content_type", "reason"])
        w.writeheader()
        w.writerows(out)
    print(f"\n{path}: {len(out)} post(s) across {len(by_brand)} brand(s)")
    print(f"Rule for the method slide: proof-device ads in rotation across creative "
          f"signatures (max {args.max_per_signature} per signature), every "
          f"campaign-platform ad, then stratified to the brand's content-type mix.")


if __name__ == "__main__":
    main()
