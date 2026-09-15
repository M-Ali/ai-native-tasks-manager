"""Build one contact sheet per brand so the creative can actually be looked at.

    python contact_sheets.py <workspace>/scan

Reads media/<handle>/*.jpg, writes contact/<handle>.jpg with each post labelled by its
shortcode, so a code in capture.xlsx can be traced back to the image it came from.

Looking at the sheets is the step that cannot be skipped. Captions describe intent;
artwork shows execution, casting, budget and production value. Half the findings in a
good scan are visible only in the image.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw

if hasattr(sys.stdout, "reconfigure"):        # Urdu / non-latin copy kills cp1252
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CELL = 300
PAD = 26


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workspace", type=Path)
    ap.add_argument("--cols", type=int, default=4)
    ap.add_argument("--selection", type=Path, default=None,
                    help="slide_selection.csv from select_for_slides.py - build sheets "
                         "from the selected posts only, so the sheet matches the deck")
    args = ap.parse_args()

    keep: dict[str, set[str]] | None = None
    if args.selection:
        import csv
        keep = {}
        with args.selection.open(encoding="utf-8", newline="") as fh:
            for r in csv.DictReader(fh):
                stem = Path(r["image_file"]).stem if r.get("image_file") else r["post_id"]
                keep.setdefault(Path(r["image_file"]).parent.name if r.get("image_file")
                                else "", set()).add(stem)
        # image_file paths are media/<handle>/<shortcode>.jpg, so the parent is the handle
        keep = {k: v for k, v in keep.items() if k}
        print(f"  selection: {sum(len(v) for v in keep.values())} post(s) "
              f"across {len(keep)} handle(s)")

    root = args.workspace / "media"
    if not root.exists():
        raise SystemExit(f"{root} not found - run collect_ig.py first")
    out = args.workspace / "contact"
    out.mkdir(parents=True, exist_ok=True)

    made = 0
    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        imgs = sorted(d.glob("*.jpg"))
        if keep is not None:
            allowed = keep.get(d.name, set())
            imgs = [p for p in imgs if p.stem in allowed] or imgs
        if not imgs:
            print(f"  {d.name}: no images, skipped")
            continue
        rows = (len(imgs) + args.cols - 1) // args.cols
        sheet = Image.new("RGB", (args.cols * CELL, rows * (CELL + PAD)), "white")
        dr = ImageDraw.Draw(sheet)
        for i, p in enumerate(imgs):
            try:
                im = Image.open(p).convert("RGB")
            except Exception as exc:
                print(f"  ! {p.name}: {type(exc).__name__}")
                continue
            im.thumbnail((CELL - 8, CELL - 8))
            x = (i % args.cols) * CELL + (CELL - im.width) // 2
            y = (i // args.cols) * (CELL + PAD) + 4
            sheet.paste(im, (x, y))
            dr.text(((i % args.cols) * CELL + 6,
                     (i // args.cols) * (CELL + PAD) + CELL + 4), p.stem, fill="black")
        f = out / f"{d.name}.jpg"
        sheet.save(f, quality=88)
        print(f"  {f}  ({len(imgs)} posts)")
        made += 1

    print(f"\n{made} contact sheet(s) in {out}.")
    print("Now READ them. Code from the artwork; use `unclear` rather than guessing.")


if __name__ == "__main__":
    main()
