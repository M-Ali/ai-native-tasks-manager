"""Set up a scan workspace: handles.csv to fill, and capture.xlsx with bound vocabularies.

    python make_handles.py brands.csv --out <workspace>/scan --category "Life insurance (Pakistan)"

brands.csv needs `brand,ring[,notes]`. Rings: client | direct | adjacent | substitute.

Writes:
    handles.csv    one row per brand, `handle` blank for you to fill by SEARCH
    capture.xlsx   Ads sheet with dropdowns, README, Brands, Codes, MarketShare

Leave `handle` blank where a brand genuinely has no account - absence is a finding and
the scripts carry it through to the output.
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

import sys

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

if hasattr(sys.stdout, "reconfigure"):        # Urdu / non-latin copy kills cp1252
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VOCAB = {
    "format": ["static", "carousel", "video", "other"],
    "audience_generation": ["genz", "geny", "genx", "boomer_plus", "mixed", "unclear"],
    "content_type": ["brand_building", "tactical_promo", "product_feature", "corporate_pr",
                     "recruitment", "calendar_topical", "csr", "ugc_repost"],
    "claim": ["protection", "savings_return", "education_marriage", "retirement",
              "faith_compliance", "national_trust", "convenience_digital",
              "price_affordability", "health", "brand_corporate", "other"],
    "register": ["fear", "aspiration", "duty", "reassurance", "pride", "humour"],
    "proof_device": ["sovereign_guarantee", "ratings_awards", "testimonial",
                     "claim_statistics", "celebrity", "expert_authority",
                     "heritage_scale", "none"],
    "who_is_in_frame": ["customer", "celebrity", "executive", "staff", "expert",
                        "none_people"],
    "production_value": ["studio", "stock", "template_graphic", "event_photo", "ugc",
                         "ai_generated", "unclear"],
    "language_lead": ["urdu_first", "english_first", "equal", "other"],
    "yes_no": ["yes", "no"],
}

COLS = ["post_id", "brand", "ring", "handle", "date", "pinned_guess", "format", "product_focus",
        "audience_generation", "content_type", "claim_primary", "claim_secondary",
        "register", "proof_device", "who_is_in_frame", "production_value",
        "ecosystem_push", "campaign_platform", "addresses_objection", "language_lead",
        "reposted_from", "source_url", "image_file", "coder", "notes"]

# column -> vocabulary key
BIND = {"format": "format", "audience_generation": "audience_generation",
        "content_type": "content_type", "claim_primary": "claim",
        "claim_secondary": "claim", "register": "register",
        "proof_device": "proof_device", "who_is_in_frame": "who_is_in_frame",
        "production_value": "production_value", "ecosystem_push": "yes_no",
        "addresses_objection": "yes_no", "language_lead": "language_lead"}

ROWS = 800


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("brands", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--category", default="<category>")
    ap.add_argument("--window", default="<state the window and stick to it>")
    ap.add_argument("--claims", default="",
                    help="comma-separated claim vocabulary for THIS category, replacing "
                         "the financial-services default. SKILL.md tells you to rewrite "
                         "it per category - this is how, so the dropdown matches your "
                         "codes instead of flagging them as validation errors.")
    args = ap.parse_args()

    out = args.out
    out.mkdir(parents=True, exist_ok=True)

    if args.claims:
        VOCAB["claim"] = [c.strip() for c in args.claims.split(",") if c.strip()]
        print(f"  claim vocabulary replaced: {len(VOCAB['claim'])} terms")

    with args.brands.open(encoding="utf-8-sig", newline="") as fh:
        brands = list(csv.DictReader(fh))
    if not brands:
        raise SystemExit(f"{args.brands}: no rows")

    handles = out / "handles.csv"
    if handles.exists():
        print(f"{handles} exists - left alone")
    else:
        with handles.open("w", encoding="utf-8", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=["brand", "ring", "handle", "notes"])
            w.writeheader()
            for b in brands:
                w.writerow({"brand": b["brand"], "ring": b.get("ring", ""),
                            "handle": "", "notes": b.get("notes", "")})
        print(f"{handles} - fill the handle column by SEARCH, never by guessing")

    wb = openpyxl.Workbook()
    ads = wb.active
    ads.title = "Ads"
    ads.append(COLS)
    for c in ads[1]:
        c.font = openpyxl.styles.Font(bold=True)
    ads.freeze_panes = "A2"
    for i, name in enumerate(COLS, start=1):
        ads.column_dimensions[get_column_letter(i)].width = \
            42 if name in ("notes", "source_url", "campaign_platform", "product_focus") else 17

    codes = wb.create_sheet("Codes")
    for i, (key, vals) in enumerate(VOCAB.items(), start=1):
        col = get_column_letter(i)
        codes[f"{col}1"] = key
        codes[f"{col}1"].font = openpyxl.styles.Font(bold=True)
        for j, v in enumerate(vals, start=2):
            codes[f"{col}{j}"] = v
        rng = f"Codes!${col}$2:${col}${len(vals)+1}"
        dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=True)
        codes.add_data_validation(dv)
        for target, vocab_key in BIND.items():
            if vocab_key == key:
                tcol = get_column_letter(COLS.index(target) + 1)
                dv.add(f"Ads!{tcol}2:{tcol}{ROWS}")

    br = wb.create_sheet("Brands")
    br.append(["brand", "ring", "handle", "followers", "notes"])
    for c in br[1]:
        c.font = openpyxl.styles.Font(bold=True)
    for b in brands:
        br.append([b["brand"], b.get("ring", ""), "", "", b.get("notes", "")])

    ms = wb.create_sheet("MarketShare")
    ms.append(["brand", "market_share_pct", "metric", "period", "source"])
    for c in ms[1]:
        c.font = openpyxl.styles.Font(bold=True)
    for b in brands:
        ms.append([b["brand"], "", "", "", ""])

    rd = wb.create_sheet("README")
    for row in [
        ["Category creative scan"],
        ["Category", args.category],
        ["Collection window", args.window],
        ["Brands in scope", len(brands)],
        ["Core category objection", "<the one objection ads are coded against>"],
        ["Row = ", "one post"],
        ["Before coding", "Read references/scan-schema.md. Code ten posts as a group, reconcile, then split."],
        ["Every row needs", "a source_url. Rows without provenance are dropped at analysis."],
        ["Code from", "the ARTWORK, not the caption. Use `unclear` rather than guessing."],
        ["Never present", "share of voice from a capped sample - every brand is equal N by construction."],
        ["Analyse with", "python analyze_scan.py capture.xlsx --out <dir>"],
    ]:
        rd.append(row)
    rd.column_dimensions["A"].width = 26
    rd.column_dimensions["B"].width = 92
    rd["A1"].font = openpyxl.styles.Font(bold=True, size=13)

    wb.move_sheet("README", offset=-4)
    path = out / "capture.xlsx"
    wb.save(path)
    print(f"{path} - {len(brands)} brand(s), {len(COLS)} columns, dropdowns bound")


if __name__ == "__main__":
    main()
