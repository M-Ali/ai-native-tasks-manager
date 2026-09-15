"""Generate a competitor-ad capture workbook with locked controlled vocabularies.

Usage:
    python make_capture_sheet.py brands.csv --out <dir> --category "Life insurance (Pakistan)"

brands.csv needs a header row: brand,ring,notes
  ring is one of: client, direct, adjacent, substitute

The workbook has five sheets:
  README       - scope, collection date, the category's core objection
  Ads          - the capture grid (dropdowns bound to Codes)
  Brands       - the brand list, with ring
  Codes        - the controlled vocabularies; edit here to adapt a category
  MarketShare  - share figures with a mandatory source column

Only openpyxl is required.
"""

from __future__ import annotations

import argparse
import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NOTE_FONT = Font(italic=True, color="808080")

RINGS = ["client", "direct", "adjacent", "substitute"]

# Controlled vocabularies. Edit the Codes sheet to adapt these per category --
# the dropdowns and analyze_audit.py both read from that sheet, not from here.
CODES: dict[str, list[str]] = {
    "channel": [
        "meta", "google_search", "google_display", "youtube", "tiktok", "linkedin",
        "press", "tv", "radio", "ooh", "website", "other",
    ],
    "format": [
        "static", "carousel", "video", "long_form", "print", "ooh", "radio", "other",
    ],
    "language": ["urdu", "english", "mixed", "regional", "other"],
    "product_line": [
        "individual_life", "group_corporate", "takaful", "savings_investment",
        "health", "child_education", "retirement", "brand_corporate", "unclear",
    ],
    "claim": [
        "protection", "savings_return", "education_marriage", "retirement",
        "faith_compliance", "national_trust", "convenience_digital",
        "price_affordability", "brand_corporate", "other",
    ],
    "register": ["fear", "aspiration", "duty", "reassurance", "pride", "humour"],
    "proof_device": [
        "sovereign_guarantee", "ratings_awards", "testimonial", "claim_statistics",
        "celebrity", "expert_authority", "heritage_scale", "none",
    ],
    "cta_channel": [
        "agent", "bank_branch", "digital_direct", "call_centre", "retail", "none",
    ],
    "audience_signal": [
        "mass", "affluent", "women", "youth", "diaspora", "sme_corporate", "rural",
    ],
    "yes_no": ["yes", "no"],
}

# (column header, code list key or None, width)
AD_COLUMNS: list[tuple[str, str | None, int]] = [
    ("ad_id", None, 8),
    ("brand", "__brands__", 28),
    ("ring", None, 12),
    ("channel", "channel", 16),
    ("date_first_seen", None, 16),
    ("date_last_seen", None, 16),
    ("format", "format", 12),
    ("duration_sec", None, 12),
    ("language", "language", 12),
    ("product_line", "product_line", 20),
    ("claim_primary", "claim", 22),
    ("claim_secondary", "claim", 22),
    ("register", "register", 14),
    ("proof_device", "proof_device", 20),
    ("cta_channel", "cta_channel", 16),
    ("audience_signal", "audience_signal", 18),
    ("addresses_objection", "yes_no", 20),
    ("variant_count", None, 14),
    ("views", None, 12),
    ("source_url", None, 46),
    ("coder", None, 10),
    ("notes", None, 46),
]

CAPTURE_ROWS = 600


def read_brands(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        raise SystemExit(f"{path} has no data rows")
    missing = {"brand", "ring"} - set(rows[0])
    if missing:
        raise SystemExit(f"{path} is missing column(s): {', '.join(sorted(missing))}")
    for row in rows:
        if row["ring"] not in RINGS:
            raise SystemExit(
                f"brand {row['brand']!r}: ring {row['ring']!r} must be one of {RINGS}"
            )
    return rows


def _header(ws, headers: list[str], widths: list[int] | None = None) -> None:
    for col, name in enumerate(headers, 1):
        cell = ws.cell(1, col, name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    if widths:
        for col, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"


def build_readme(wb: Workbook, category: str, brand_count: int) -> None:
    ws = wb.create_sheet("README")
    rows = [
        ("Competitor communications audit", ""),
        ("Category", category),
        ("Collection started", date.today().isoformat()),
        ("Collection window", "<e.g. Sep 2025 - Aug 2026 -- state it and stick to it>"),
        ("Brands in scope", str(brand_count)),
        ("", ""),
        ("Core category objection", "<the single objection ads are coded against, e.g. "
                                    "'will they actually pay out?'>"),
        ("Row = concept or asset?", "<decide once, apply consistently; note variant_count>"),
        ("", ""),
        ("Before coding", "Read references/coding-schema.md. Code the same ten ads as a "
                          "group, reconcile, then split the work."),
        ("Every row needs", "a source_url. Rows without provenance are dropped at analysis."),
        ("Screenshots", "<folder path -- ads come down; keep the evidence>"),
        ("", ""),
        ("Analyse with", "python analyze_audit.py capture.xlsx --out <dir>"),
    ]
    for r, (k, v) in enumerate(rows, 1):
        ws.cell(r, 1, k).font = Font(bold=True)
        cell = ws.cell(r, 2, v)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if v.startswith("<"):
            cell.font = NOTE_FONT
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 82


def build_codes(wb: Workbook) -> None:
    ws = wb.create_sheet("Codes")
    for col, (name, values) in enumerate(CODES.items(), 1):
        cell = ws.cell(1, col, name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        ws.column_dimensions[get_column_letter(col)].width = max(len(name), 22)
        for r, value in enumerate(values, 2):
            ws.cell(r, col, value)
    ws.freeze_panes = "A2"


def build_brands(wb: Workbook, brands: list[dict[str, str]]) -> None:
    ws = wb.create_sheet("Brands")
    _header(ws, ["brand", "ring", "notes"], [34, 14, 60])
    for r, row in enumerate(brands, 2):
        ws.cell(r, 1, row["brand"])
        ws.cell(r, 2, row["ring"])
        ws.cell(r, 3, row.get("notes", "")).alignment = Alignment(wrap_text=True)


def build_market_share(wb: Workbook, brands: list[dict[str, str]]) -> None:
    ws = wb.create_sheet("MarketShare")
    _header(
        ws,
        ["brand", "metric", "value", "period", "source", "source_url"],
        [34, 22, 12, 14, 44, 44],
    )
    ws.cell(2, 2, "<e.g. gross written premium share, %>").font = NOTE_FONT
    for r, row in enumerate(brands, 2):
        ws.cell(r, 1, row["brand"])
    note = ws.cell(len(brands) + 3, 1, "Every figure needs a citable source. Do not enter a "
                                       "number you cannot attribute.")
    note.font = NOTE_FONT


def build_ads(wb: Workbook, brands: list[dict[str, str]]) -> None:
    ws = wb.create_sheet("Ads", 0)
    _header(ws, [c[0] for c in AD_COLUMNS], [c[2] for c in AD_COLUMNS])

    code_cols = {name: get_column_letter(i) for i, name in enumerate(CODES, 1)}

    for idx, (name, code_key, _width) in enumerate(AD_COLUMNS, 1):
        letter = get_column_letter(idx)
        if code_key is None:
            continue
        if code_key == "__brands__":
            last = len(brands) + 1
            formula = f"=Brands!$A$2:$A${last}"
        else:
            letter_src = code_cols[code_key]
            last = len(CODES[code_key]) + 1
            formula = f"=Codes!${letter_src}$2:${letter_src}${last}"
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
        dv.error = f"Pick a value from the {code_key} list"
        dv.errorTitle = "Not in the controlled vocabulary"
        ws.add_data_validation(dv)
        dv.add(f"{letter}2:{letter}{CAPTURE_ROWS + 1}")

    # ad_id and ring fill themselves so coders cannot desynchronise them.
    brand_col = get_column_letter(2)
    for r in range(2, CAPTURE_ROWS + 2):
        ws.cell(r, 1, f'=IF({brand_col}{r}="","",ROW()-1)')
        ws.cell(r, 3, f'=IFERROR(VLOOKUP({brand_col}{r},Brands!$A:$B,2,FALSE),"")')
    ws.auto_filter.ref = f"A1:{get_column_letter(len(AD_COLUMNS))}{CAPTURE_ROWS + 1}"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("brands", type=Path, help="CSV with columns brand,ring,notes")
    ap.add_argument("--out", type=Path, required=True, help="output directory")
    ap.add_argument("--category", default="<category>", help="category name for the README")
    args = ap.parse_args()

    brands = read_brands(args.brands)
    args.out.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)
    build_ads(wb, brands)
    build_readme(wb, args.category, len(brands))
    build_brands(wb, brands)
    build_codes(wb)
    build_market_share(wb, brands)

    path = args.out / "capture.xlsx"
    wb.save(path)
    print(f"Wrote {path} -- {len(brands)} brands, {CAPTURE_ROWS} capture rows")
    print("Next: fill the README scope lines, then read references/coding-schema.md before coding.")


if __name__ == "__main__":
    main()
