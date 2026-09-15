"""Migrate the SLIC audit from the competitor-comms-audit schema to the nine-dimension
category-creative-scan schema, folding in enriched dates and captions.

    uv run --with openpyxl python workspace/audit/migrate_to_scan_schema.py

Reads  workspace/audit/capture.xlsx        (old 22-col schema, 144 coded rows)
       workspace/audit/posts.csv           (shortcode -> image_file, reposted_from)
       workspace/audit/enriched/posts.csv  (shortcode -> real date, full caption)
Writes workspace/audit/capture_scan.xlsx   (new 24-col schema)

Carries over every code that maps one-for-one. The five new dimensions are left blank
for coding from the contact sheets - they are judgement, and a migration script has no
business guessing them.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import openpyxl

AUDIT = Path(__file__).parent
SKILL = Path.home() / ".claude" / "skills" / "category-creative-scan" / "scripts"
sys.path.insert(0, str(SKILL))
from make_handles import COLS as NEW_COLS  # noqa: E402  the schema, single-sourced


def read_csv(path: Path, key: str) -> dict[str, dict]:
    if not path.exists():
        print(f"  {path} not found - skipping that merge")
        return {}
    with path.open(encoding="utf-8", newline="") as fh:
        return {r[key]: r for r in csv.DictReader(fh) if r.get(key)}


def main() -> None:
    old_path = AUDIT / "capture.xlsx"
    wb_old = openpyxl.load_workbook(old_path, data_only=True)
    ws_old = wb_old["Ads"]
    hdr = [c.value for c in ws_old[1]]
    old_rows = [dict(zip(hdr, [c.value for c in r])) for r in ws_old.iter_rows(min_row=2)
                if r[1].value]
    print(f"  {len(old_rows)} row(s) in the old sheet")

    posts = read_csv(AUDIT / "posts.csv", "shortcode")
    rich = read_csv(AUDIT / "enriched" / "posts.csv", "shortcode")
    print(f"  {len(posts)} post record(s), {len(rich)} enriched record(s)")

    def shortcode(url: str | None) -> str:
        if not url:
            return ""
        parts = [p for p in str(url).split("/") if p]
        return parts[-1] if parts else ""

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ads"
    ws.append(NEW_COLS)
    for c in ws[1]:
        c.font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "A2"

    dated = captioned = 0
    for o in old_rows:
        sc = shortcode(o.get("source_url"))
        p = posts.get(sc, {})
        r = rich.get(sc, {})
        date = r.get("date") or o.get("date_last_seen") or ""
        if r.get("date"):
            dated += 1
        cap = r.get("caption") or ""
        if cap:
            captioned += 1
        note = (o.get("notes") or "")
        if cap:
            note = (note + f"  CAPTION: {cap}")[:1200]

        ws.append([
            sc,                                   # post_id
            o.get("brand"), o.get("ring") or p.get("ring", ""), p.get("handle", ""),
            date, o.get("format"),
            "",                                   # product_focus        - to code
            "",                                   # audience_generation  - to code
            "",                                   # content_type         - to code
            o.get("claim_primary"), o.get("claim_secondary"),
            o.get("register"), o.get("proof_device"),
            "",                                   # who_is_in_frame      - to code
            "",                                   # production_value     - to code
            "",                                   # ecosystem_push       - to code
            "",                                   # campaign_platform    - to code
            o.get("addresses_objection"),
            "mixed",                              # language_lead - refine when coding
            p.get("reposted_from", ""),
            o.get("source_url"), p.get("image_file", ""),
            o.get("coder"), note,
        ])

    for name in ("Codes", "Brands", "MarketShare", "README"):
        if name in wb_old.sheetnames:
            src = wb_old[name]
            dst = wb.create_sheet(name)
            for row in src.iter_rows(values_only=True):
                dst.append(row)

    out = AUDIT / "capture_scan.xlsx"
    wb.save(out)
    print(f"\n{out}")
    print(f"  {len(old_rows)} rows migrated | {dated} with a real post date | "
          f"{captioned} with a full caption")
    print("  5 new dimensions left blank - code them from contact/ , do not guess.")


if __name__ == "__main__":
    main()
