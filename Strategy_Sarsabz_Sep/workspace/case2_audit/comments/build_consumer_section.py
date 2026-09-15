"""Consumer insight section for the Sarsabz Case 2 deck, built from the comment-analysis outputs.

    uv run --with python-pptx python build_consumer_section.py --out <deck.pptx> [--base <deck.pptx>] [--cover] [--front]

Slides: (cover) · divider · the job · the fear · the names · what farmers seek.

Two guards, because these slides get quoted:
- Numbers are read live from theme_counts.csv, brand_counts.csv and coding_report.md, so the slides
  cannot drift from the coded file.
- Every quote is checked against comments.csv before anything is built. A quote not found verbatim
  stops the build unless it is marked as a translation.

--base copies an existing deck (never edits it); --front moves the new slides to the start.
The output path is never overwritten: a clash becomes -v2, -v3.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(__file__).parent
NAVY = RGBColor(0x10, 0x2A, 0x43)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x74, 0x80)
ACCENT = RGBColor(0xC8, 0x10, 0x2E)
PALE = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SKY = RGBColor(0x9F, 0xB4, 0xC7)
BARGREY = RGBColor(0xB8, 0xC2, 0xCC)

# key: (text as written, note, is_translation)
QUOTES = {
    "recipe": ("Waqt kasht: 1 bag dap + enrich First pani: 1 bag nitrophas+ 1 bag zabardat urea Second pani: 1 bag "
               "nitrophas plus 1 bag can gawara+3kg sulphur Third pani: 25 kg soluble potash+ 20kg can gawara... "
               "Ye fertilizer plan kaisa ha?",
               "District Muzaffargarh - a whole season's plan, posted for strangers to approve.", False),
    "lockup": ("Zabrdast urea ma zink ha aur np ma phosphorus ha to akthy dalen gy to zink phosphorus to fix ho jay ge",
               "Zabardast has zinc, NP has phosphorus - together they lock up.", False),
    "fatima": ("Fatima fertilizer walon ne NP+CAN+ZINC SULFATE mila ke chatta kiya hai result bhe behar hai aor aap "
               "bol rahai in ko aapus main na milaain",
               "Fatima's own field practice, against the creator the farmer is watching.", False),
    "rule": ("app ne old videos me bataya hy k pehly pani pher can gawara dy do aur neche DAP khaad hota hy ,ab keh "
             "rhy ho k in dono ko mix na karo",
             "One rule in the old videos, the opposite now.", False),
    "npdap": ("How many bags of Sarsabz Nitrophos give the same result as one DAP, or can it not replace DAP?",
              "Translated from Urdu - the switching question, asked outright.", True),
    "canyield": ("Can gawara say paidawar mien izafa hota hy kia",
                 "The only yield question in the file: does CAN raise yield at all?", False),
    "which": ("fawara khad kis company ki achi ha SAR sbz ya Fatima wali",
              "Which is better - Sarsabz or the Fatima one? They are the same company.", False),
    "mafia": ("Medicine mafia se paise lekar video banner rehta",
              "Farmers suspect creators take company money.", False),
    "bioactive": ("Zabardast Urea me mojood zinc bio active form me hoti hai",
                  "A farmer-facing reply explaining Engro's product by name.", False),
}

THEMES = [("mixing_compatibility", "What mixes with what"), ("timing_dose", "Timing and dose"),
          ("application_method", "Flood or broadcast"), ("soil_type", "Which soil"),
          ("yield_result", "Yields and results"), ("price_cost", "Price and cost"),
          ("region_request", "Advice for my region"), ("contact_reply_request", "Reply to me / your number"),
          ("np_vs_dap_substitution", "NP instead of DAP")]

NAMES = [("Urea (generic)", "Urea", "product"), ("CAN / Gawara (product)", "CAN (gawara)", "product"),
         ("NP / Nitrophos (product)", "NP (Nitrophos)", "product"), ("DAP (product)", "DAP", "product"),
         ("Engro Zabardast", "Engro Zabardast", "brand"), ("FFC Sona", "FFC Sona", "brand"),
         ("Sarsabz", "Sarsabz", "client"), ("Fatima", "Fatima", "brand"), ("Pak Arab", "Pak Arab", "brand")]


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()


def check_quotes() -> None:
    corpus = norm("\n".join(r["text"] for r in csv.DictReader((HERE / "comments.csv").open(encoding="utf-8"))))
    bad = []
    for key, (text, _note, translated) in QUOTES.items():
        if translated:
            continue
        parts = [p.strip(" .") for p in re.split(r"\.\.\.|…", text) if len(p.strip(" .")) >= 6]
        if not parts or not all(norm(p) in corpus for p in parts):
            bad.append(key)
    if bad:
        raise SystemExit("Refusing to build - quotes not found verbatim in comments.csv: " + ", ".join(bad))


def load_counts():
    themes = {r["theme"]: int(r["comments"]) for r in csv.DictReader((HERE / "theme_counts.csv").open(encoding="utf-8"))}
    brands = {r["brand"]: int(r["mentions"]) for r in csv.DictReader((HERE / "brand_counts.csv").open(encoding="utf-8"))}
    report = (HERE / "coding_report.md").read_text(encoding="utf-8")
    n = int(re.search(r"Comments coded: \*\*(\d+)\*\*", report).group(1))
    cov = re.search(r"Matched at least one theme: \*\*([\d,]+ \([\d.]+%\))\*\*", report).group(1)
    return themes, brands, n, cov


def txt(slide, l, t, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, space=6, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(line)
        p.alignment = align
        p.space_after = Pt(space)
        f = p.runs[0].font
        f.size, f.bold, f.italic, f.color.rgb, f.name = Pt(size), bold, italic, color, "Segoe UI"
    return tb


def rect(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def frame(prs, kicker, title, conclusion=None, source=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 1.0, NAVY)
    txt(s, 0.6, 0.12, 12.1, 0.3, kicker, 10, True, SKY, space=0)
    txt(s, 0.6, 0.42, 12.1, 0.5, title, 22, True, WHITE, space=0)
    if conclusion:
        rect(s, 0.6, 6.35, 12.1, 0.62, NAVY)
        txt(s, 0.85, 6.45, 11.6, 0.5, conclusion, 12, True, WHITE, space=0)
    if source:
        txt(s, 0.6, 7.05, 12.1, 0.35, source, 8, False, MUTED, space=0)
    return s


def quote_tile(s, x, y, w, key, size=11):
    text, note, translated = QUOTES[key]
    rect(s, x, y, 0.045, 1.25, ACCENT)
    body = f"“{text}”"
    txt(s, x + 0.22, y, w - 0.22, 0.95, body, size, False, INK, space=0, italic=True)
    txt(s, x + 0.22, y + 0.98, w - 0.22, 0.3, note, 9.5, False, MUTED, space=0)


def build(prs, cover: bool):
    themes, brands, n, cov = load_counts()
    source = (f"Source: {n:,} YouTube comments on independent agronomy channels, supplied by the client team; "
              f"themes matched {cov}. Counts rank where to read - findings come from reading every comment. "
              f"Self-selected corpus, not a measure of what most farmers think. Translations labelled.")

    if cover:
        s = prs.slides.add_slide(prs.slide_layouts[6])
        rect(s, 0, 0, 13.333, 7.5, NAVY)
        rect(s, 0.9, 2.0, 1.6, 0.06, ACCENT)
        txt(s, 0.9, 2.3, 11.4, 1.0, "Sarsabz", 48, True, WHITE)
        txt(s, 0.9, 3.35, 11.4, 0.6, "Case 2  ·  Brand audit and the Big Idea", 22, False, SKY)
        txt(s, 0.9, 4.1, 11.4, 0.5,
            "1  What farmers say    2  How the category talks    3  The Brand Meaning Ladder    4  The Big Idea    5  The campaign",
            13, False, WHITE)
        txt(s, 0.9, 6.4, 11.4, 0.4,
            "Fatima Fertilizer creative agency pitch 2026  ·  Working draft, 15 September 2026", 12, False, SKY)

    # divider
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    rect(s, 0.9, 2.3, 1.6, 0.06, ACCENT)
    txt(s, 0.9, 2.6, 11.4, 0.9, "What farmers say", 40, True, WHITE)
    txt(s, 0.9, 3.55, 11.4, 1.0,
        [f"{n:,} comments on independent agronomy channels, every one read.",
         "The consumer pulse behind the Big Idea."], 18, False, SKY, space=4)

    # the job
    s = frame(prs, "WHAT FARMERS SAY  ·  THE JOB", "What farmers are trying to solve",
              "The largest subject in the file is not whether a fertilizer works. It is how to use it without waste.",
              source)
    rect(s, 0.6, 1.35, 0.05, 2.6, ACCENT)
    txt(s, 0.85, 1.35, 5.9, 2.2, f"“{QUOTES['recipe'][0]}”", 15, False, INK, space=0, italic=True)
    txt(s, 0.85, 3.65, 5.9, 0.3, QUOTES["recipe"][1], 10, False, MUTED, space=0)
    txt(s, 0.6, 4.4, 6.2, 1.6,
        "The job is not choosing a brand. It is getting the combination, the timing and the method right, "
        "so nothing they paid for is wasted.", 16, True, NAVY, space=0)
    txt(s, 7.3, 1.35, 5.4, 0.3, "WHAT THEY ASK ABOUT  ·  COMMENTS", 10, True, MUTED, space=0)
    top = max(themes.get(k, 0) for k, _ in THEMES) or 1
    y = 1.75
    for key, label in THEMES:
        v = themes.get(key, 0)
        txt(s, 7.3, y, 2.4, 0.3, label, 10.5, False, INK, space=0)
        rect(s, 9.75, y + 0.05, max(0.05, 2.3 * v / top), 0.22, NAVY if key == "mixing_compatibility" else BARGREY)
        txt(s, 12.1, y, 0.6, 0.3, str(v), 10.5, True, NAVY, PP_ALIGN.RIGHT, space=0)
        y += 0.5

    # the fear
    s = frame(prs, "WHAT FARMERS SAY  ·  THE PAIN", "The fear is waste, not failure",
              "Nobody questions whether NP and CAN raise yield. They ask whether mixing will waste them - and get "
              "contradictory answers, including from Fatima's own field practice.", source)
    for i, key in enumerate(["lockup", "npdap", "fatima", "canyield", "rule", "mafia"]):
        quote_tile(s, 0.6 + (i % 2) * 6.2, 1.35 + (i // 2) * 1.62, 5.9, key)

    # the names
    s = frame(prs, "WHAT FARMERS SAY  ·  THE BRAND", "Farmers name the product, not the brand",
              "Sarsabz makes the products farmers discuss most, and is almost never named. Engro shows a "
              "fertilizer brand can be asked for by name.", source)
    txt(s, 0.6, 1.35, 6.4, 0.3, "COMMENTS NAMING EACH PRODUCT OR BRAND", 10, True, MUTED, space=0)
    peak = max(brands.get(k, 0) for k, _, _ in NAMES) or 1
    y = 1.75
    for key, label, kind in NAMES:
        v = brands.get(key, 0)
        colour = ACCENT if kind == "client" else (BARGREY if kind == "product" else NAVY)
        txt(s, 0.6, y, 2.1, 0.3, label, 11, kind == "client", ACCENT if kind == "client" else INK, space=0)
        rect(s, 2.75, y + 0.05, max(0.05, 3.6 * v / peak), 0.24, colour)
        txt(s, 6.4, y, 0.6, 0.3, str(v), 11, True, ACCENT if kind == "client" else NAVY, PP_ALIGN.RIGHT, space=0)
        y += 0.47
    txt(s, 0.6, 6.0, 6.4, 0.3, "Grey = products  ·  Navy = brands  ·  Red = Sarsabz", 9, False, MUTED, space=0)
    rect(s, 7.4, 1.35, 5.3, 4.8, PALE)
    txt(s, 7.65, 1.55, 4.8, 0.3, "WHAT IT MEANS", 10, True, MUTED, space=0)
    txt(s, 7.65, 1.95, 4.8, 1.6,
        [f"CAN {brands.get('CAN / Gawara (product)', 0)} and NP {brands.get('NP / Nitrophos (product)', 0)} "
         f"mentions - Sarsabz {brands.get('Sarsabz', 0)}. Farmers buy Fatima's products as generics.",
         f"Engro Zabardast ({brands.get('Engro Zabardast', 0)}) is the only fertilizer brand named routinely - "
         "and people explain it:"], 12, False, INK, space=8)
    txt(s, 7.65, 3.55, 4.8, 0.6, f"“{QUOTES['bioactive'][0]}”", 11, False, INK, space=0, italic=True)
    rect(s, 7.65, 4.35, 0.045, 1.0, ACCENT)
    txt(s, 7.85, 4.35, 4.6, 0.6, f"“{QUOTES['which'][0]}”", 11, False, INK, space=0, italic=True)
    txt(s, 7.85, 4.95, 4.6, 0.5, QUOTES["which"][1], 9.5, False, MUTED, space=0)

    # what they seek
    s = frame(prs, "WHAT FARMERS SAY  ·  WHAT THEY SEEK", "What farmers want, and what it means for Sarsabz",
              "Everything farmers seek is proof and instruction. None of it is a salute.", source)
    rows = [
        ("A straight answer on what mixes with what",
         "One NP and CAN compatibility and timing guide, signed by the maker - today field staff and creators contradict each other."),
        ("A plan for their own soil, crop and district",
         "Proof and advice by soil and region. A national claim does not answer 'my land is sandy'."),
        ("Evidence, not opinion (translated: 'is there any written proof?')",
         "Farmers ask for proof unprompted. Sarsabz holds 500+ demo plots and shows none on screen."),
        ("Someone who replies",
         "Farmers ask creators for replies and WhatsApp numbers - an audience waiting for the Sarsabz Pakistan App."),
        ("Advice that is not paid for",
         "Sponsored creator content starts with a credibility deficit: disclose every partnership and attach evidence."),
    ]
    txt(s, 0.6, 1.3, 4.6, 0.3, "FARMERS SEEK", 10, True, MUTED, space=0)
    txt(s, 5.5, 1.3, 7.2, 0.3, "WHAT IT MEANS FOR SARSABZ", 10, True, MUTED, space=0)
    y = 1.65
    for i, (seek, means) in enumerate(rows):
        if i % 2 == 0:
            rect(s, 0.6, y - 0.05, 12.1, 0.9, PALE)
        txt(s, 0.75, y + 0.08, 4.5, 0.75, seek, 12, True, NAVY, space=0)
        txt(s, 5.5, y + 0.08, 7.0, 0.75, means, 11.5, False, INK, space=0)
        y += 0.9


def free_path(p: Path) -> Path:
    if not p.exists():
        return p
    n = 2
    while (q := p.with_name(f"{p.stem}-v{n}{p.suffix}")).exists():
        n += 1
    return q


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--base", type=Path)
    ap.add_argument("--cover", action="store_true")
    ap.add_argument("--front", action="store_true", help="move the new slides to the start of the deck")
    args = ap.parse_args()

    check_quotes()
    if args.base:
        prs = Presentation(str(args.base))
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    before = len(prs.slides)
    build(prs, args.cover)
    added = len(prs.slides) - before
    if args.front and before:
        lst = prs.slides._sldIdLst
        new = list(lst)[-added:]
        for i, el in enumerate(new):
            lst.remove(el)
            lst.insert(i, el)
    out = free_path(args.out)
    prs.save(out)
    print(f"{out}  ({len(prs.slides)} slides; consumer section = {added})")


if __name__ == "__main__":
    main()
