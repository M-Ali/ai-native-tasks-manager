"""Render one target-group profile card (a single 16:9 slide) from a JSON spec.

    python build_tg_card.py spec.json [--base deck.pptx]

Writes the spec's `out` path, never overwriting: a clash becomes -v2, -v3...
With --base, the card is appended to a copy of an existing deck.

THE GUARDS, AND WHY

A persona slide is the most quoted and least checked page in any deck: its numbers get repeated for years.
So this refuses to build when the card would mislead -

- a pain point or trigger with no `evidence`      an unevidenced line briefs people on a guess
- a "%" in media or interests with no `source`    invented media percentages are the genre's standard fiction
- an interests block whose source names no population and no total
                                                  an index describes the population that produced it
- text that does not fit its box                  a card that overflows in PowerPoint was never really checked

Spec shape: see references/spec-example.json.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NAVY = RGBColor(0x10, 0x2A, 0x43)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x74, 0x80)
ACCENT = RGBColor(0xC8, 0x10, 0x2E)
PALE = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SKY = RGBColor(0x9F, 0xB4, 0xC7)

def fits(text: str, w: float, h: float, size: float) -> bool:
    """Estimate whether text fits, by line rather than by area.

    Segoe UI averages about 0.0062in of width and 0.019in of line height per point of size, so a 13pt
    line in a 9in box holds roughly 110 characters. An area-based estimate gets short wide boxes badly
    wrong - it rejects a one-line subtitle that fits easily - so the budget is lines x chars-per-line.
    """
    chars_per_line = max(1, w / (0.0062 * size))
    lines = max(1, h / (0.019 * size))
    # 0.9, not 1.0: every wrapped line wastes part of its last row, and a card that overflows in
    # PowerPoint after passing this check is worse than one that made you shorten a line.
    return len(text) <= chars_per_line * lines * 0.9


def check(spec: dict) -> None:
    errors = []
    for key in ("pain_points", "triggers"):
        for item in spec.get(key, []):
            if not item.get("evidence"):
                errors.append(f"{key}: {item.get('title', '?')!r} has no evidence - "
                              "every line on a persona card has to be traceable")
    for key in ("media", "interests"):
        block = spec.get(key) or {}
        body = json.dumps(block.get("items", []), ensure_ascii=False)
        if "%" in body and not block.get("source"):
            errors.append(f"{key}: a percentage appears with no source. Cite the survey, or drop the number "
                          "and describe the channel instead")
    ints = spec.get("interests") or {}
    if ints.get("items"):
        src = ints.get("source", "")
        if not src:
            errors.append("interests: no source. An affinity index describes the population that produced it")
        elif not re.search(r"\d", src):
            errors.append(f"interests: source {src!r} names no population size or date - "
                          "say whose data this is and how many people are in it")
    if not spec.get("sources"):
        errors.append("no `sources` line: the card has to say where it came from")
    if errors:
        raise SystemExit("Refusing to build:\n  - " + "\n  - ".join(errors))
    for key in ("pain_points", "triggers"):
        if len(spec.get(key, [])) > 6:
            print(f"warning: {len(spec[key])} {key}; about 6 fit the card")


def txt(slide, l, t, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, space=5, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(line)
        p.alignment = align
        p.space_after = Pt(space)
        f = p.runs[0].font
        f.size, f.bold, f.italic, f.color.rgb, f.name = Pt(size), bold, italic, color, "Segoe UI"
    body = " ".join(str(x) for x in lines)
    if not fits(body, w, h, size):
        raise SystemExit(f"Refusing to build: text overflows its box ({len(body)} chars in {w:.1f}x{h:.1f}in "
                         f"at {size}pt). Shorten it or give it more room:\n    {body[:120]}...")
    return tb


def rect(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def panel(slide, l, t, w, h, title, lines, size=11, accent=ACCENT):
    rect(slide, l, t, w, h, PALE)
    rect(slide, l, t, 0.06, h, accent)
    txt(slide, l + 0.22, t + 0.14, w - 0.4, 0.3, title.upper(), 10, True, accent, space=0)
    txt(slide, l + 0.22, t + 0.52, w - 0.4, h - 0.66, lines, size, False, INK, space=6)


def build(prs, spec: dict) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 1.05, NAVY)
    txt(s, 0.45, 0.14, 9.0, 0.45, spec["name"], 30, True, WHITE, space=0)
    txt(s, 0.45, 0.63, 9.0, 0.3, spec.get("subtitle", ""), 13, False, SKY, space=0)
    if spec.get("tag"):
        rect(s, 10.9, 0.22, 2.0, 0.36, ACCENT)
        txt(s, 10.9, 0.29, 2.0, 0.25, spec["tag"].upper(), 10, True, WHITE, align=PP_ALIGN.CENTER, space=0)

    # left column: household, interests (or the gap), identity facts
    y = 1.25
    if spec.get("household"):
        panel(s, 0.45, y, 4.0, 2.4, "Who he is", spec["household"], 10.5)
        y += 2.55
    ints = spec.get("interests") or {}
    items = ints.get("items") or []
    lines = [f"{i['label']}  -  i{i['index']}  ({i['users']:,} users)" if isinstance(i, dict) else str(i)
             for i in items] or [ints.get("gap", "No interest-affinity dataset covers this audience.")]
    if ints.get("source"):
        lines = lines + [ints["source"]]
    panel(s, 0.45, y, 4.0, 6.35 - y, ints.get("title", "Interests and affinity index"), lines, 10.5)

    # centre column: mindset, pain points
    panel(s, 4.65, 1.25, 4.6, 1.9, "Behavioural mindset", spec["mindset"], 11.5)
    pains = [f"{p['title']}: {p['detail']}" for p in spec.get("pain_points", [])]
    panel(s, 4.65, 3.3, 4.6, 3.05, f"Pain points ({len(pains)})", pains, 10.5)

    # right column: triggers, and what would fill the gaps
    trig = [f"{t['title']}: {t['detail']}" for t in spec.get("triggers", [])]
    panel(s, 9.45, 1.25, 3.45, 2.4, "Buying triggers", trig, 10)
    media = spec.get("media") or {}
    mlines = [f"{m['channel']}: {m['note']}" if isinstance(m, dict) else str(m) for m in media.get("items", [])]
    if media.get("source"):
        mlines = mlines + [media["source"]]
    panel(s, 9.45, 3.8, 3.45, 2.55, media.get("title", "Where to reach him"), mlines, 10)

    if spec.get("conclusion"):
        rect(s, 0.45, 6.5, 12.45, 0.55, NAVY)
        txt(s, 0.7, 6.61, 12.0, 0.35, spec["conclusion"], 12, True, WHITE, space=0)
    # two lines of room: a card's sources list is long by design, and it must not be cut off
    txt(s, 0.45, 7.05, 12.45, 0.42, "Sources: " + "  ·  ".join(spec["sources"]), 8, False, MUTED, space=0)


def free_path(p: Path) -> Path:
    if not p.exists():
        return p
    n = 2
    while (q := p.with_name(f"{p.stem}-v{n}{p.suffix}")).exists():
        n += 1
    return q


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", type=Path)
    ap.add_argument("--base", type=Path, help="existing deck to append to (copied, never edited)")
    args = ap.parse_args()
    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    check(spec)

    if args.base:
        tmp = args.base.with_suffix(".tmp.pptx")
        shutil.copy(args.base, tmp)
        prs = Presentation(str(tmp))
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    build(prs, spec)
    out = Path(spec["out"])
    out = out if out.is_absolute() else args.spec.parent / out
    out.parent.mkdir(parents=True, exist_ok=True)
    dest = free_path(out)
    prs.save(dest)
    if args.base:
        tmp.unlink(missing_ok=True)
    print(f"{dest}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
