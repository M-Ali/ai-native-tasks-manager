"""Shape the coded YouTube read into the inputs category-creative-scan's build_deck.py expects.

    uv run --with openpyxl --with pillow python deck_prep.py

Writes scan/deck/: capture.xlsx (sheet "Ads", the skill's COLS), handles.csv, profiles.csv and
contact/<handle>.jpg. reads.yaml is authored by hand in the same folder.

The skill's collector and capture sheet were built for Instagram. The coded data here is owned
YouTube, so this script maps it into the same columns rather than changing the skill. `format` is
"video" for every row; `pinned_guess` is "no" (YouTube has no pinned grid).

Contact sheet rule, stated because selection shapes what the audience sees: each brand's 25 most-viewed
uploads in the window - i.e. where the brand put its media weight - labelled with views and platform.
"""
import csv
from pathlib import Path

import openpyxl
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
OUT = HERE / "deck"
(OUT / "contact").mkdir(parents=True, exist_ok=True)

COLS = ["post_id", "brand", "ring", "handle", "date", "pinned_guess", "format", "product_focus",
        "audience_generation", "content_type", "claim_primary", "claim_secondary",
        "register", "proof_device", "who_is_in_frame", "production_value",
        "ecosystem_push", "campaign_platform", "addresses_objection", "language_lead",
        "reposted_from", "source_url", "image_file", "coder", "notes"]

BRANDS = [
    # brand, ring, handle, codes file, thumbs dir, audience shown on the slide
    ("Sarsabz", "client", "Sarsabz",
     HERE.parent / "owned" / "codes_window.csv", HERE.parent / "owned" / "thumbs", "396K YouTube"),
    ("Engro Fertilizers", "direct", "engrofertilizerbrands",
     HERE / "engro" / "codes_window.csv", HERE / "engro" / "thumbs", "81.2K YouTube"),
]
PER_SHEET, COLS_GRID, TW, TH, LABEL = 25, 5, 320, 180, 26


def views_label(v: int) -> str:
    return f"{v / 1e6:.1f}M" if v >= 1e6 else (f"{v / 1e3:.0f}K" if v >= 1e3 else str(v))


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Ads"
ws.append(COLS)
font = ImageFont.truetype("arial.ttf", 15)

with (OUT / "handles.csv").open("w", newline="", encoding="utf-8") as hf, \
        (OUT / "profiles.csv").open("w", newline="", encoding="utf-8") as pf:
    hw = csv.writer(hf)
    hw.writerow(["brand", "ring", "handle"])
    pw = csv.writer(pf)
    pw.writerow(["brand", "handle", "followers"])
    for brand, ring, handle, codes, thumbs, audience in BRANDS:
        rows = list(csv.DictReader(codes.open(encoding="utf-8")))
        hw.writerow([brand, ring, handle])
        pw.writerow([brand, handle, audience])
        for r in rows:
            rec = dict.fromkeys(COLS, "")
            rec.update(post_id=r["video_id"], brand=brand, ring=ring, handle=handle, date=r["published"],
                       pinned_guess="no", format="video", product_focus=r["product_focus"],
                       audience_generation=r["audience_generation"], content_type=r["content_type"],
                       claim_primary=r["claim_primary"], register=r["register"],
                       proof_device=r["proof_device"], who_is_in_frame=r["who_is_in_frame"],
                       production_value=r["production_value"], ecosystem_push=r["ecosystem_push"],
                       campaign_platform=r["platform"], addresses_objection=r["addresses_objection"],
                       source_url=r["url"], image_file=f"thumbs/{r['video_id']}.jpg",
                       coder=r["coder"], notes=r["notes"])
            ws.append([rec[c] for c in COLS])

        top = sorted(rows, key=lambda r: -int(r["views"]))[:PER_SHEET]
        nrows = (len(top) + COLS_GRID - 1) // COLS_GRID
        sheet = Image.new("RGB", (COLS_GRID * TW + (COLS_GRID + 1) * 6, nrows * (TH + LABEL) + 6), "white")
        d = ImageDraw.Draw(sheet)
        for i, r in enumerate(top):
            x = 6 + (i % COLS_GRID) * (TW + 6)
            y = 6 + (i // COLS_GRID) * (TH + LABEL)
            try:
                im = Image.open(thumbs / f"{r['video_id']}.jpg").convert("RGB")
                im = im.resize((TW, TH))
                sheet.paste(im, (x, y))
            except Exception:
                d.rectangle([x, y, x + TW, y + TH], fill=(230, 230, 230))
            d.text((x + 2, y + TH + 4), f"{views_label(int(r['views']))}  ·  {r['platform'][:26]}",
                   fill=(40, 40, 40), font=font)
        sheet.save(OUT / "contact" / f"{handle}.jpg", quality=88)
        print(f"{brand}: {len(rows)} rows, contact sheet of {len(top)}")

wb.save(OUT / "capture.xlsx")
print("wrote", OUT)
