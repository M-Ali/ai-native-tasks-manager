"""Code claim / register / proof_device / objection across the 288-post deep read.

    uv run --with openpyxl python workspace/audit/build_claims_deep.py

Two sources, in priority order:

1. **The original 144-post coding**, carried over by post_id. A post coded from the
   artwork at 144 must not be re-coded differently at 288 - the deeper read adds posts,
   it does not revise judgements already made from the image.
2. **Rules over the caption** for the 144 posts the first pass never saw.

`proof_device` carries the headline number in this analysis, so its rules are written
tightly and default to `none` - which is both the honest default and the one that makes
the finding harder, not easier, to claim.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import openpyxl

AUDIT = Path(__file__).parent
sys.path.insert(0, str(AUDIT))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from build_codes import norm  # noqa: E402

FIELDS = ["claim_primary", "claim_secondary", "register", "proof_device",
          "addresses_objection"]


def K(c1, reg, proof="none", obj="no", c2="other"):
    return dict(zip(FIELDS, [c1, c2, reg, proof, obj]))


# (caption needle, codes). Brand-agnostic where the pattern is; brand-scoped where not.
GLOBAL: list[tuple[str, dict]] = [
    # --- proof: a performance number. The rarest and most important code in the set.
    ("claims paid",            K("national_trust", "reassurance", "claim_statistics", "yes", "protection")),
    ("billion in claims",      K("national_trust", "reassurance", "claim_statistics", "yes", "protection")),
    ("complaints resolved",    K("national_trust", "reassurance", "claim_statistics", "yes", "other")),
    ("e-kachehri",             K("national_trust", "reassurance", "claim_statistics", "yes", "other")),
    ("claim your matured",     K("national_trust", "reassurance", "claim_statistics", "yes", "savings_return")),
    # --- proof: a real customer, named
    ("policyholder",           K("national_trust", "reassurance", "testimonial", "yes", "protection")),
    ("sara rizvi",             K("national_trust", "reassurance", "testimonial", "yes", "protection")),
    ("mahin khan",             K("national_trust", "reassurance", "testimonial", "yes", "protection")),
    ("hear sahibzadi",         K("national_trust", "reassurance", "testimonial", "yes", "protection")),
    ("mantahaa",               K("health", "aspiration", "testimonial", "no", "other")),
    ("saima aasim",            K("health", "aspiration", "testimonial", "no", "other")),
    # --- proof: heritage and scale
    ("since 1972",             K("national_trust", "pride", "heritage_scale", "no", "protection")),
    ("five decades",           K("national_trust", "pride", "heritage_scale", "no", "protection")),
    ("for over 50 years",      K("national_trust", "pride", "heritage_scale", "no", "protection")),
    ("trusted by your parents", K("national_trust", "reassurance", "heritage_scale", "yes", "protection")),
    ("ten years is more",      K("brand_corporate", "pride", "heritage_scale")),
    ("half a million downloads", K("convenience_digital", "pride", "heritage_scale")),
    # --- objection handling without a proof device
    ("think again",            K("protection", "reassurance", "none", "yes")),
    ("common belief",          K("protection", "reassurance", "none", "yes")),
    ("myth",                   K("protection", "reassurance", "none", "yes")),
    ("takaful is different",   K("faith_compliance", "duty", "none", "yes")),
    ("tie your camel",         K("faith_compliance", "duty", "expert_authority", "yes")),
    ("halal way",              K("faith_compliance", "duty", "none", "yes")),
    ("always more time",       K("protection", "fear", "none", "yes")),
    ("career has an end date", K("retirement", "fear", "none", "yes", "savings_return")),
    # --- ordinary claims
    ("discount",               K("price_affordability", "aspiration", "none", "no", "convenience_digital")),
    ("% off",                  K("price_affordability", "aspiration", "none", "no", "convenience_digital")),
    ("earn rewards",           K("price_affordability", "aspiration")),
    ("retirement",             K("retirement", "reassurance", "none", "no", "savings_return")),
    ("pension",                K("retirement", "reassurance", "none", "no", "savings_return")),
    ("education",              K("education_marriage", "duty", "none", "no", "protection")),
    ("school fees",            K("education_marriage", "duty", "none", "no", "protection")),
    ("child",                  K("education_marriage", "duty", "none", "no", "protection")),
    ("shariah",                K("faith_compliance", "reassurance", "none", "no", "savings_return")),
    ("takaful",                K("faith_compliance", "reassurance", "none", "no", "protection")),
    ("halal",                  K("faith_compliance", "reassurance", "none", "no", "savings_return")),
    ("globewell",              K("protection", "aspiration", "none", "no", "health")),
    ("worldwide",              K("protection", "reassurance", "none", "no", "health")),
    ("sehat zindagi",          K("price_affordability", "reassurance", "none", "no", "health")),
    ("independence day",       K("national_trust", "pride")),
    ("azaadi",                 K("national_trust", "pride")),
    ("eid",                    K("faith_compliance", "reassurance")),
    ("ramadan",                K("faith_compliance", "reassurance")),
    ("muharram",               K("faith_compliance", "reassurance")),
    ("hiring",                 K("brand_corporate", "aspiration")),
    ("we are hiring",          K("brand_corporate", "aspiration")),
    ("career opportunit",      K("brand_corporate", "aspiration")),
    ("conference",             K("brand_corporate", "pride", "expert_authority")),
    ("podcast",                K("faith_compliance", "reassurance", "expert_authority")),
    ("episode",                K("faith_compliance", "reassurance", "expert_authority")),
    ("app",                    K("convenience_digital", "aspiration")),
    ("whatsapp",               K("convenience_digital", "reassurance")),
    ("raast",                  K("convenience_digital", "reassurance")),
    ("savings",                K("savings_return", "aspiration")),
    ("invest",                 K("savings_return", "aspiration")),
    ("wellbeing",              K("health", "aspiration")),
    ("health",                 K("health", "reassurance")),
]

# Pak-Qatar puts an AA / AM2+ ratings badge on essentially every creative.
BRAND_PROOF = {"Pak-Qatar Family Takaful": "ratings_awards"}

DEFAULT = K("brand_corporate", "reassurance")


def main() -> None:
    prior = {}
    scan = AUDIT / "capture_scan.xlsx"
    if scan.exists():
        wb = openpyxl.load_workbook(scan, data_only=True)
        ws = wb["Ads"]
        hdr = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            d = dict(zip(hdr, r))
            if d.get("post_id") and d.get("claim_primary"):
                prior[str(d["post_id"])] = {f: d.get(f) for f in FIELDS}
        print(f"  carried over {len(prior)} post(s) coded from the artwork at 144")

    rows = list(csv.DictReader((AUDIT / "deep" / "posts.csv").open(encoding="utf-8")))
    out, reused, ruled, defaulted = [], 0, 0, 0
    for r in rows:
        pid = r["shortcode"]
        if pid in prior:
            out.append({"post_id": pid, **prior[pid]})
            reused += 1
            continue
        cap = norm(r.get("caption") or "")
        codes = None
        if cap:
            for needle, c in GLOBAL:
                if needle in cap:
                    codes = dict(c)
                    break
        if codes:
            ruled += 1
        else:
            codes = dict(DEFAULT)
            defaulted += 1
        if codes["proof_device"] == "none" and r["brand"] in BRAND_PROOF:
            codes["proof_device"] = BRAND_PROOF[r["brand"]]
        out.append({"post_id": pid, **codes})

    path = AUDIT / "codes_claims_deep.csv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["post_id"] + FIELDS)
        w.writeheader()
        w.writerows(out)
    print(f"{path}: {len(out)} rows")
    print(f"  {reused} carried over from the artwork-coded 144")
    print(f"  {ruled} coded by caption rule")
    print(f"  {defaulted} fell to the default (brand_corporate / reassurance / no proof)")


if __name__ == "__main__":
    main()
