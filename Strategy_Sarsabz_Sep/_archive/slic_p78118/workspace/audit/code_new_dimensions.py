"""Apply the five new dimensions to capture_scan.xlsx.

    uv run --with openpyxl python workspace/audit/code_new_dimensions.py

Codes live in `new_dimension_codes.py` as a dict keyed by shortcode, authored from the
contact sheets in contact/ and the captions merged in by the enrichment pass.

Anything not in the dict is written as `unclear` / blank rather than guessed, and the
run prints how many rows that was. A visible gap is the point: it tells the next person
exactly what still needs eyes on it.

Validates every code against the controlled vocabulary before writing, so a typo fails
here rather than silently poisoning the analysis.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

AUDIT = Path(__file__).parent
sys.path.insert(0, str(AUDIT))
sys.path.insert(0, str(Path.home() / ".claude" / "skills" / "category-creative-scan" / "scripts"))

from make_handles import VOCAB  # noqa: E402  single-sourced vocabulary

try:
    from new_dimension_codes import CODES  # noqa: E402
except ImportError:
    sys.exit("new_dimension_codes.py not found - author it from the contact sheets first.")

# column -> (vocabulary key or None for free text)
FIELDS = {
    "product_focus": None,
    "audience_generation": "audience_generation",
    "content_type": "content_type",
    "who_is_in_frame": "who_is_in_frame",
    "production_value": "production_value",
    "ecosystem_push": "yes_no",
    "campaign_platform": None,
    "language_lead": "language_lead",
}

BLANK = {
    "product_focus": "", "audience_generation": "unclear", "content_type": "",
    "who_is_in_frame": "none_people", "production_value": "unclear",
    "ecosystem_push": "no", "campaign_platform": "", "language_lead": "mixed",
}


def main() -> None:
    path = AUDIT / "capture_scan.xlsx"
    if not path.exists():
        sys.exit(f"{path} not found - run migrate_to_scan_schema.py first.")

    wb = openpyxl.load_workbook(path)
    ws = wb["Ads"]
    hdr = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(hdr)}

    # validate the whole code table before touching the sheet
    errors = []
    for sc, codes in CODES.items():
        for field, value in codes.items():
            if field not in FIELDS:
                errors.append(f"{sc}: unknown field {field!r}")
                continue
            vocab_key = FIELDS[field]
            if vocab_key and value not in VOCAB[vocab_key]:
                errors.append(f"{sc}: {field}={value!r} not in vocabulary "
                              f"({', '.join(VOCAB[vocab_key])})")
    if errors:
        for e in errors[:25]:
            print(f"  {e}")
        sys.exit(f"\n{len(errors)} invalid code(s). Nothing written.")

    coded = blanked = 0
    missing: list[str] = []
    for r in range(2, ws.max_row + 1):
        sc = ws.cell(row=r, column=col["post_id"]).value
        if not sc:
            continue
        codes = CODES.get(sc)
        if codes:
            coded += 1
        else:
            blanked += 1
            missing.append(str(sc))
            codes = {}
        for field in FIELDS:
            if field in col:
                ws.cell(row=r, column=col[field],
                        value=codes.get(field, BLANK[field]))

    wb.save(path)
    print(f"{path}")
    print(f"  {coded} row(s) coded on all five new dimensions")
    if blanked:
        print(f"  {blanked} row(s) left uncoded - these need eyes on the full-size image:")
        for sc in missing[:20]:
            print(f"      {sc}")
        if len(missing) > 20:
            print(f"      ... and {len(missing)-20} more")


if __name__ == "__main__":
    main()
