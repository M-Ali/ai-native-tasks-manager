"""Turn a filled capture workbook into findings.

Usage:
    python analyze_audit.py capture.xlsx --out <dir>

Writes findings.md plus CSVs of every table, so the design team can rebuild the
charts in the deck without retyping numbers.

Produces:
  - sample frame (what was actually looked at)
  - claim x register crosstab -- the territory map data
  - per-brand claim profile
  - share-of-voice proxy vs share of market
  - objection-handling and proof-device rates
  - coder drift check

Only openpyxl is required. No spend is estimated anywhere; presence is counted
and labelled as a proxy.
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

DROP_NO_SOURCE = True


def load_rows(path: Path) -> tuple[list[dict], list[dict], dict[str, str]]:
    wb = load_workbook(path, data_only=True)

    def sheet_rows(name: str) -> list[dict]:
        if name not in wb.sheetnames:
            return []
        ws = wb[name]
        headers = [c.value for c in ws[1]]
        out = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            record = {h: v for h, v in zip(headers, row) if h}
            if any(v not in (None, "") for v in record.values()):
                out.append(record)
        return out

    ads = [r for r in sheet_rows("Ads") if r.get("brand")]
    share = [r for r in sheet_rows("MarketShare") if r.get("value") not in (None, "")]
    rings = {r["brand"]: r.get("ring", "") for r in sheet_rows("Brands") if r.get("brand")}
    return ads, share, rings


def write_csv(path: Path, headers: list[str], rows: list[list]) -> None:
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(headers)
        w.writerows(rows)


def pct(n: int, total: int) -> str:
    return f"{(100 * n / total):.0f}%" if total else "-"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    ads, share, rings = load_rows(args.workbook)
    args.out.mkdir(parents=True, exist_ok=True)

    dropped = []
    if DROP_NO_SOURCE:
        kept = []
        for row in ads:
            (kept if row.get("source_url") else dropped).append(row)
        ads = kept

    if not ads:
        raise SystemExit("No coded ads with a source_url found -- nothing to analyse.")

    total = len(ads)
    for row in ads:
        row["ring"] = row.get("ring") or rings.get(row["brand"], "")

    L: list[str] = []
    add = L.append

    add("# Competitor communications audit - findings")
    add("")
    add(f"Source workbook: `{args.workbook.name}`")
    add("")

    # --- 1. sample frame -----------------------------------------------------
    brands = Counter(r["brand"] for r in ads)
    channels = Counter(r.get("channel") or "unspecified" for r in ads)
    dates = [r.get("date_first_seen") for r in ads if r.get("date_first_seen")]

    add("## 1. What we looked at")
    add("")
    add(f"**{total} ads** coded across **{len(brands)} brands**.")
    if dates:
        add(f"Earliest ad dated {min(dates)}; latest {max(dates)}.")
    if dropped:
        add(f"{len(dropped)} row(s) excluded for having no `source_url`.")
    add("")
    add("| Brand | Ring | Ads | Share of coded ads |")
    add("|---|---|---:|---:|")
    for brand, n in brands.most_common():
        add(f"| {brand} | {rings.get(brand, '')} | {n} | {pct(n, total)} |")
    add("")
    add("| Channel | Ads |")
    add("|---|---:|")
    for ch, n in channels.most_common():
        add(f"| {ch} | {n} |")
    add("")
    write_csv(
        args.out / "sample_frame.csv",
        ["brand", "ring", "ads", "share_pct"],
        [[b, rings.get(b, ""), n, round(100 * n / total, 1)] for b, n in brands.most_common()],
    )

    # --- 2. territory map ----------------------------------------------------
    claims = sorted({r.get("claim_primary") for r in ads if r.get("claim_primary")})
    registers = sorted({r.get("register") for r in ads if r.get("register")})
    grid: dict[tuple[str, str], int] = Counter(
        (r.get("claim_primary"), r.get("register"))
        for r in ads
        if r.get("claim_primary") and r.get("register")
    )

    add("## 2. How the category talks - territory map")
    add("")
    add("Claim territory (rows) against emotional register (columns). Empty cells are the")
    add("white space; the concept should aim at one.")
    add("")
    add("| Claim \\ Register | " + " | ".join(registers) + " | Total |")
    add("|---" * (len(registers) + 2) + "|")
    for claim in claims:
        cells = [grid.get((claim, reg), 0) for reg in registers]
        add(f"| {claim} | " + " | ".join(str(c) if c else "-" for c in cells) + f" | {sum(cells)} |")
    add("")
    write_csv(
        args.out / "territory_map.csv",
        ["claim_primary", *registers],
        [[c, *[grid.get((c, r), 0) for r in registers]] for c in claims],
    )

    empty = [(c, r) for c in claims for r in registers if grid.get((c, r), 0) == 0]
    if empty:
        add(f"**{len(empty)} of {len(claims) * len(registers)} claim/register combinations "
            f"are unoccupied.**")
        add("")
        add("Interrogate these before concepting - some are empty because they do not work, but")
        add("the defensible territory is usually among them.")
        add("")

    # --- 3. per-brand claim profile -----------------------------------------
    by_brand: dict[str, Counter] = defaultdict(Counter)
    for r in ads:
        if r.get("claim_primary"):
            by_brand[r["brand"]][r["claim_primary"]] += 1

    add("## 3. Per-brand claim profile")
    add("")
    add("| Brand | " + " | ".join(claims) + " |")
    add("|---" * (len(claims) + 1) + "|")
    for brand in brands:
        row = by_brand[brand]
        add(f"| {brand} | " + " | ".join(str(row.get(c, 0) or "-") for c in claims) + " |")
    add("")
    write_csv(
        args.out / "brand_claim_profile.csv",
        ["brand", *claims],
        [[b, *[by_brand[b].get(c, 0) for c in claims]] for b in brands],
    )

    # --- 4. share of voice proxy vs share of market -------------------------
    add("## 4. Who is heard - share of voice proxy vs share of market")
    add("")
    add("**Share of voice here is a presence proxy** - counted ads, and views where the")
    add("platform publishes them. It is not spend, and must not be presented as spend.")
    add("")
    views = defaultdict(int)
    for r in ads:
        v = r.get("views")
        if isinstance(v, (int, float)):
            views[r["brand"]] += int(v)
    share_by_brand = {r["brand"]: r for r in share if r.get("brand")}

    add("| Brand | Ads | SOV proxy | Reported views | Market share | Source |")
    add("|---|---:|---:|---:|---:|---|")
    for brand, n in brands.most_common():
        ms = share_by_brand.get(brand, {})
        ms_val = ms.get("value")
        ms_txt = f"{ms_val}" if ms_val not in (None, "") else "-"
        add(
            f"| {brand} | {n} | {pct(n, total)} | {views.get(brand, 0) or '-'} | "
            f"{ms_txt} | {ms.get('source', '-') or '-'} |"
        )
    add("")
    if not share_by_brand:
        add("> No market-share figures entered. Fill the `MarketShare` sheet with cited")
        add("> regulator or annual-report figures - this comparison is usually the single")
        add("> most useful chart in the section.")
        add("")
    write_csv(
        args.out / "sov_vs_som.csv",
        ["brand", "ads", "sov_proxy_pct", "views", "market_share", "source"],
        [
            [
                b, n, round(100 * n / total, 1), views.get(b, 0),
                share_by_brand.get(b, {}).get("value", ""),
                share_by_brand.get(b, {}).get("source", ""),
            ]
            for b, n in brands.most_common()
        ],
    )

    # --- 5. proof and objection ---------------------------------------------
    proofs = Counter(r.get("proof_device") or "unspecified" for r in ads)
    objection = Counter(str(r.get("addresses_objection") or "unspecified").lower() for r in ads)
    ctas = Counter(r.get("cta_channel") or "unspecified" for r in ads)
    langs = Counter(r.get("language") or "unspecified" for r in ads)

    add("## 5. Proof, objection handling, distribution and language")
    add("")
    add("| Proof device | Ads | Share |")
    add("|---|---:|---:|")
    for k, n in proofs.most_common():
        add(f"| {k} | {n} | {pct(n, total)} |")
    add("")
    yes = objection.get("yes", 0)
    add(f"**Ads addressing the category's core objection head-on: {yes} of {total} "
        f"({pct(yes, total)}).**")
    if total and yes / total < 0.25:
        add("")
        add("Under a quarter of category advertising tackles the core objection directly.")
        add("If the objection is what actually blocks purchase, that is the gap - and")
        add("addressing it is available to whoever moves first.")
    add("")
    add("| CTA channel | Ads | | Language | Ads |")
    add("|---|---:|---|---|---:|")
    cta_rows, lang_rows = ctas.most_common(), langs.most_common()
    for i in range(max(len(cta_rows), len(lang_rows))):
        c = f"{cta_rows[i][0]} | {cta_rows[i][1]}" if i < len(cta_rows) else " | "
        g = f"{lang_rows[i][0]} | {lang_rows[i][1]}" if i < len(lang_rows) else " | "
        add(f"| {c} | | {g} |")
    add("")

    # --- 6. coder drift ------------------------------------------------------
    coders = sorted({r.get("coder") for r in ads if r.get("coder")})
    if len(coders) > 1:
        add("## 6. Coder drift check")
        add("")
        add("Each coder's distribution across claim territories. A profile that diverges")
        add("sharply from the others is drift, not insight - recode before analysing.")
        add("")
        add("| Coder | Ads | " + " | ".join(claims) + " |")
        add("|---|---:|" + "---:|" * len(claims))
        for coder in coders:
            rows = [r for r in ads if r.get("coder") == coder]
            counts = Counter(r.get("claim_primary") for r in rows)
            add(
                f"| {coder} | {len(rows)} | "
                + " | ".join(pct(counts.get(c, 0), len(rows)) for c in claims)
                + " |"
            )
        add("")

    add("## Next")
    add("")
    add("The tables above are inputs, not the finding. Write one sentence naming the vacant")
    add("territory and one paragraph proving it from this evidence, then state what the client")
    add("can credibly own - anchored in an asset competitors structurally cannot copy.")
    add("")

    path = args.out / "findings.md"
    path.write_text("\n".join(L), encoding="utf-8")
    print(f"Wrote {path}")
    print(f"       + sample_frame.csv, territory_map.csv, brand_claim_profile.csv, sov_vs_som.csv")
    print(f"Analysed {total} ads across {len(brands)} brands.")
    if dropped:
        print(f"Excluded {len(dropped)} row(s) with no source_url.")


if __name__ == "__main__":
    main()
