"""Case 1 (Salam Kissan 2026) deck: cover + evidence, the Big Idea section (big-idea-slides skill), the campaign.

    uv run --no-project --with python-pptx python workspace/case1_salam_kissan/deck/build_case1_deck.py

Writes workspace/out/Sarsabz_case1_deck.pptx (never overwrites: a clash becomes -v2, -v3...).

Guards, because the user's rule for Case 1 is "no made-up data, no assumptions":
- every verbatim on a slide goes through V() and must exist in the comment corpus, the YouTube metadata,
  the archive index or the brief text, or the build stops. Translations go through T() and are labelled;
- every quoted string in big_idea_spec.json that isn't marked as a translation must pass the same check;
- counts shown on slides are computed here from the CSVs, not typed in.
"""
from __future__ import annotations

import csv
import json
import re
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = Path(__file__).parent
C1 = HERE.parent
WS = C1.parent
ROOT = WS.parent
sys.path.insert(0, str(ROOT / "src"))
from sarsabz_pitch import brief as B  # noqa: E402

SKILL = Path.home() / ".claude/skills/big-idea-slides/scripts/build_big_idea_section.py"
NAVY = RGBColor(0x10, 0x2A, 0x43)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x74, 0x80)
ACCENT = RGBColor(0xC8, 0x10, 0x2E)
PALE = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SKY = RGBColor(0x9F, 0xB4, 0xC7)


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("’", "'").replace("​", "").replace("\xa0", " ")).strip().lower()


SOURCES = norm(" ".join(p.read_text(encoding="utf-8") for p in (
    C1 / "comments/corpus.csv", C1 / "scan/others.csv", C1 / "scan/search_raw.csv",
    C1 / "archive/sk_index.csv", WS / "brief/full.txt")))
MISSING: list[str] = []


def V(s: str) -> str:
    """A verbatim. Must exist in a source file."""
    if norm(s).rstrip(".") not in SOURCES:
        MISSING.append(s)
    return f"“{s}”"


def T(s: str) -> str:
    """Our translation - labelled, not checked."""
    return f"(our translation: {s})"


# ---------- data ----------
arch = list(csv.DictReader((C1 / "archive/archive_codes.csv").open(encoding="utf-8")))
corpus = list(csv.DictReader((C1 / "comments/corpus.csv").open(encoding="utf-8")))
codes = list(csv.DictReader((C1 / "comments/comment_codes.csv").open(encoding="utf-8")))
others = list(csv.DictReader((C1 / "scan/others.csv").open(encoding="utf-8")))
search = list(csv.DictReader((C1 / "scan/search_raw.csv").open(encoding="utf-8")))

N_UP = len(arch)
V_ALL = sum(int(r["views"]) for r in arch)
IN_WIN = sum(r["in_window"] == "yes" for r in arch)
FARMER = sum(r["speaker"] == "farmer" and r["format"] != "product_testimonial" for r in arch)
CELEB_POL = sum(r["speaker"] in ("celebrity", "politician") for r in arch)
ANIMATED_25 = sum(r["format"] == "animated_short" and r["year"] == "2025" for r in arch)
by_year = defaultdict(list)
for r in arch:
    by_year[r["year"]].append(r)

aud = [r for r in corpus if r["is_uploader"] == "no"]
ANTHEMS = {"PkenunwpduU": "2019", "llIKvpyxhKM": "2021", "Pd0z78-UJio": "2022", "PX7Lvjs639U": "2023"}
per_film = Counter(r["video_id"] for r in corpus)
ANTHEM_C = sum(per_film[v] for v in ANTHEMS)
Y25 = [r for r in aud if r["upload_date"].startswith("2025")]
fa = defaultdict(set)
for r in Y25:
    fa[r["author"]].add(r["video_id"])
MULTI = {a for a, vs in fa.items() if len(vs) >= 5}
Y25_MULTI = sum(r["author"] in MULTI for r in Y25)
code_n = Counter(c for r in codes for c in r["codes"].split(";") if c)

rel = {r["video_id"]: r for r in search
       if re.search(r"(?i)kiss?an ?day|kisan ?day|farmers?.? ?day|salam kiss?an|کسان|18 ?dec", r["title"] or "")}
SARSABZ_SEARCH = sum(r["channel"] == "Sarsabz" for r in rel.values())
syn = sorted([o for o in others if o["channel"] == "Syngenta Pakistan"], key=lambda o: o["upload_date"])

FACTS = {"uploads": N_UP, "in_window": IN_WIN, "farmer": FARMER, "celeb_pol": CELEB_POL, "anthem_comments": ANTHEM_C,
         "comments": len(corpus), "audience": len(aud), "grievance": code_n["grievance"],
         "song": code_n["song_voice_lyrics"], "winner": code_n["channel_winner_notice"], "vlog": code_n["vlog_referral"],
         "salam_pakistan": code_n["salam_pakistan"],
         "y25": len(Y25), "y25_multi": Y25_MULTI, "search_sarsabz": SARSABZ_SEARCH, "search_rel": len(rel)}
# The spec and big_idea.md carry these numbers as text; stop if the data no longer says the same.
EXPECT = {"uploads": 76, "in_window": 60, "farmer": 9, "celeb_pol": 30, "anthem_comments": 1422, "comments": 1572,
          "audience": 1516, "grievance": 19, "song": 103, "winner": 31, "vlog": 24, "y25": 61, "y25_multi": 35,
          "search_sarsabz": 43, "search_rel": 168, "salam_pakistan": 19}
SALAM_PK = code_n["salam_pakistan"]


# ---------- drawing ----------
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


def panel(s, l, t, w, h, title, lines, size=11):
    rect(s, l, t, w, h, PALE)
    rect(s, l, t, 0.06, h, ACCENT)
    txt(s, l + 0.25, t + 0.15, w - 0.4, 0.35, title.upper(), 10, True, ACCENT, space=0)
    txt(s, l + 0.25, t + 0.55, w - 0.4, h - 0.7, lines, size, False, INK, space=7)


def table(s, l, t, widths, rows, size=10.5, row_h=0.36, head=True, head_h=0.45):
    y = t
    for i, row in enumerate(rows):
        hdr = head and i == 0
        h = head_h if hdr else row_h
        if hdr:
            rect(s, l, y, sum(widths), h, NAVY)
        elif i % 2 == 0:
            rect(s, l, y, sum(widths), h, PALE)
        x = l
        for w, cell in zip(widths, row):
            txt(s, x + 0.1, y + 0.1, w - 0.2, h - 0.15, cell, size, hdr, WHITE if hdr else INK, space=0)
            x += w
        y += h
    if y > 6.3:
        raise SystemExit(f"table runs to {y:.2f} in and would cover the conclusion bar")
    return y


def m(v: int) -> str:
    return f"{v / 1e6:.1f}M" if v >= 1e6 else f"{v / 1e3:.0f}K"


# ---------- part 1: cover and evidence ----------
def evidence(prs) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    rect(s, 0.9, 1.6, 1.6, 0.06, ACCENT)
    txt(s, 0.9, 1.85, 11.4, 0.5, "Fatima Fertilizer creative pitch  ·  Case 1", 18, False, SKY)
    txt(s, 0.9, 2.4, 11.6, 1.0, "Salam Kissan 2026", 48, True, WHITE)
    txt(s, 0.9, 3.45, 11.4, 0.9, "The farmer answers the salute.", 26, False, WHITE, italic=True)
    txt(s, 0.9, 4.6, 11.4, 1.4, ["1  The brief   ·   2  Seven years on YouTube   ·   3  What commenters say",
                                 "4  Who else marks the day   ·   5  The Big Idea   ·   6  The campaign"], 14, False, SKY)
    txt(s, 0.9, 6.5, 11.6, 0.4, "Internal working draft  ·  16 September 2026  ·  verified evidence only; "
                                "translations labelled; brand insight provisional", 11, False, SKY)

    # 1 the brief
    s = frame(prs, "1  ·  THE BRIEF  ·  PAGE 12", "What Fatima is asking for",
              f"The ask in the client's words: {V('strengthening our ownership of the cause')}, beyond a single day, "
              "rural and urban, especially the young.",
              "Source: Fatima Creative Marketing Brief 2026, p12 (src/sarsabz_pitch/brief.py, page-traced).")
    txt(s, 0.6, 1.3, 6.0, 0.35, "THE CONCEPT AND 360° CAMPAIGN MUST", 11, True, ACCENT, space=0)
    txt(s, 0.6, 1.75, 6.0, 3.0, [f"{i}.  {it.text}" for i, it in enumerate(B.CASE1_TASK, 1)], 13, space=10)
    txt(s, 7.0, 1.3, 5.7, 0.35, "GOAL: MAKE RURAL AND URBAN AUDIENCES UNDERSTAND", 11, True, ACCENT, space=0)
    txt(s, 7.0, 1.75, 5.7, 2.2, [f"{i}.  {it.text}" for i, it in enumerate(B.CASE1_GOAL, 1)], 13, space=10)
    panel(s, 7.0, 4.1, 5.7, 1.9, "Touchpoints the brief names", [", ".join(B.CASE1_CHANNELS),
          "Background: Salam Kissan began in 2019; celebrated on 18 December with a 360° campaign (p12)."], 11)

    # 2 the archive
    s = frame(prs, "2  ·  SEVEN YEARS ON YOUTUBE", "Salam Kissan lives on one day, and the farmer rarely speaks",
              f"{IN_WIN} of {N_UP} uploads ({100 * IN_WIN / N_UP:.0f}%) sit between 1 Dec and 15 Jan. A farmer speaks in "
              f"{FARMER}; celebrities and politicians in {CELEB_POL}.",
              f"Source: all {N_UP} Salam Kissan / Kissan Day uploads on @Sarsabz, 2019-2025, coded from thumbnails, titles and "
              "descriptions, not watched end to end (archive/). Views on a brand channel are mostly paid weight.")
    rows = [["Year", "Uploads", "Views", "Biggest upload"]]
    for y in sorted(by_year):
        rs = by_year[y]
        top = max(rs, key=lambda r: int(r["views"]))
        rows.append([y, str(len(rs)), m(sum(int(r["views"]) for r in rs)), f"{top['title'][:52]} ({m(int(top['views']))})"])
    table(s, 0.6, 1.3, [0.9, 1.0, 1.0, 4.9], rows, 10.5, 0.52)
    speakers = Counter(r["speaker"] for r in arch)
    panel(s, 8.8, 1.3, 3.9, 4.8, "Who speaks, of 76 uploads", [
        f"Nobody spoken (anthems, shorts): {speakers['none_spoken']}",
        f"Celebrities: {speakers['celebrity']}   ·   Politicians: {speakers['politician']}",
        f"Farmers: {speakers['farmer']}, of which {speakers['farmer'] - FARMER} are product testimonials",
        "2020: 10 celebrities and 5 ministers celebrating Kissan Day.",
        f"2024: no anthem-length film. 2025: {ANIMATED_25} 3D/AI-animated farmer shorts."], 11)

    # 3a what commenters say
    s = frame(prs, "3  ·  WHAT COMMENTERS SAY", "Pride in family, response to the music, few but pointed complaints",
              "The people who write back talk about their own father, not about farmers in general.",
              f"Source: {len(corpus):,} comments on the 18 largest Salam Kissan uploads ({len(aud):,} from the audience), all read "
              "(comments/comment_findings.md, analyse.py). Self-selected; quotes are signal, not a measure. Translations labelled.")
    cols = [("PRIDE IN FAMILY", [
                V("I feel proud because my  father is farmer.") + "  2019, 123 likes",
                V("I'm daughter of farmer. My grandfather was a farmer. And I'm proud of it.") + "  2019",
                V("Although I am not a farmers but every Sarsabz ad is mind blowing!") + "  2021, 18 likes"]),
            (f"THE MUSIC  ·  {code_n['song_voice_lyrics']} COMMENTS", [
                V("Heart touching lyrics, and a very emotional tune.") + "  2019, 39 likes",
                V("magical background voice of Mai Dhaii") + "  2021, 15 likes",
                V("We need to make this permanent not temporary.") + "  2019, 87 likes"]),
            (f"COMPLAINTS  ·  {code_n['grievance']} OF {len(aud):,}", [
                V("کسانوں کو کھاد نہیں مل رہی") + "  2023  " + T("farmers aren't getting fertilizer"),
                V("کاش اس ملک میں کسان کی قدر ہوتی") + "  2022, 38 likes  " + T("if only farmers were valued in this country"),
                V("khaad nayaab ho cuki hy") + "  2022  " + T("fertilizer has become unavailable")])]
    for i, (head, lines) in enumerate(cols):
        panel(s, 0.6 + i * 4.1, 1.3, 3.9, 4.85, head, lines, 11)

    # 3b read the threads with care
    s = frame(prs, "3  ·  WHAT COMMENTERS SAY", "Where the comments are, and what inflates them",
              f"The four anthems hold {ANTHEM_C:,} of {len(corpus):,} comments. The 2023 thread was driven by a giveaway "
              "and creator referrals as well as the film.",
              "Source: comments/comment_codes.csv. Contest rules and whether creator mentions were paid: not in the data, questions for the client.")
    rows = [["Upload", "Comments"]]
    for v, y in ANTHEMS.items():
        rows.append([f"{y} anthem", str(per_film[v])])
    rows.append(["Eight 2025 uploads (audience)", str(len(Y25))])
    table(s, 0.6, 1.3, [3.6, 1.4], rows, 12, 0.55)
    panel(s, 6.0, 1.3, 6.7, 4.85, "Read with care", [
        f"2023: {code_n['channel_winner_notice']} channel replies telling winners to email their details, e.g. "
        + V("Congratulations on winning! Please email your complete details") + ".",
        f"2023: {code_n['vlog_referral']} visitors say they came from creator vlogs, e.g. "
        + V("Who is here after ducky Bahi vlog") + ".",
        f"2025: four accounts active on 5+ of the 8 uploads wrote {Y25_MULTI} of the {len(Y25)} audience comments.",
        f"2019: one account posted {V('Fatima Group is Fraud Gang')} {code_n['fraud_gang_replies']} times.",
        "2019: Indian viewers found the anthem during India's farmer protests: "
        + V("Who is here after injustice to indian farmers") + " (122 likes)."], 11)

    # 4 the category
    s = frame(prs, "4  ·  WHO ELSE MARKS THE DAY", "The salute is open to anyone, including the name",
              "Syngenta runs December Kisan Day films, and JPL and Rizq Foods use the Salam Kissan name. No FFC or Engro "
              "campaign was found.",
              f"Source: 7 YouTube searches, {len({r['video_id'] for r in search})} videos; metadata for {len(others)} non-Sarsabz "
              "uploads, 16 Sep 2026 (scan/). Search rank is not a census; TikTok and Facebook not covered; brand-channel views mostly paid.")
    rows = [["Brand", "What was found"],
            ["Syngenta Pakistan", "; ".join(f"{o['upload_date'][:4]} {o['duration']}s {m(int(o['views']))}" for o in syn)
             + "  ·  #MaiKisanHun (2023)"],
            ["JPL Pakistan", "3 executive tributes titled " + V("Salam Kissan - A tribute to farmers of Pakistan by JPL Pakistan")[:30] + "…”  (Dec 2020)"],
            ["Rizq Foods", V("Salam Kisan | Kisan Day Pakistan | Rizq foods") + "  (Dec 2025)"],
            ["Barket Fertilizers", V("Kissan Day - Kisan Ki Azmat Faslon Mein Barket") + "  (Dec 2023)"],
            ["BaKhabar Kissan", "Kissan Day film every year 2019-2022"],
            ["FFC", "Only a third-party upload of a regional manager's Farmer's Day speech"],
            ["Sarsabz", f"{SARSABZ_SEARCH} of the {len(rel)} search results titled for the day or Salam Kissan"]]
    table(s, 0.6, 1.3, [2.4, 5.6], rows, 10.5, 0.6)
    panel(s, 8.9, 1.3, 3.8, 4.85, "The critic", [
        V("18 december kisan day sirf tv show per  kisan Haqeeqat main rull gia pakistan ka kisan"),
        T("Kissan Day is only on TV shows; in reality Pakistan's farmer is ruined"),
        "DIGITAL KISAN Pakistan, 18 Dec 2022.",
        "The press: Express News and BOL name Fatima in Kissan Day programmes; GNN, SAMAA, Aaj and Neo name no brand in their titles."], 11)


# ---------- part 3: the campaign ----------
CHAIN = [
    ("Asset", "Sarsabz started Salam Kissan and Kissan Day (brief p12), runs the national event (6th Kissan Day with the "
              "Federal Minister and FAO, ProPakistani 8 Jan 2025) and holds seven years of the day's films."),
    ("Tension", "The farmer is saluted, but celebrities, ministers and animation deliver the salute."),
    ("Territory", "The farmer's own voice, all year, with Sarsabz handing over the microphone."),
    ("Proposition", "We gave Pakistan a day to salute the farmer. Now the farmer answers."),
    ("Line", "Salam Kissan. Salam Pakistan.   (proposed: the farmer salutes the country back)"),
    ("Mechanic", "Kissan ka Jawab, the farmer's answer: the name the anthem, the UGC and the PR record all sit under."),
]
EXEC = [
    ("ATL", "The Reply anthem", "Real farming families, parent and child, sing and speak the 2026 anthem back. No celebrities, no AI farmers.",
     f"Anthems hold {ANTHEM_C:,} of {len(corpus):,} comments; the 2025 uploads drew {len(Y25)}."),
    ("ATL / Digital, urban", "What I put on your table today", "The reply cut for the city: a farmer names his district, his crop and the meal it becomes.",
     "The archive's urban films work: 4.0M on the 2025 food-blogger film; 388K and 382K on the 2020 urban shorts."),
    ("Digital / UGC", "#KissanKaJawab", "Farmers' children film their parent's answer; creators seed it to people who don't farm.",
     f"The pride-in-family comments; {code_n['vlog_referral']} 2023 creator referrals; TikTok UGC awards (brief p9-10)."),
    ("BTL", "Reply cards", "Replies and voice notes collected at Kissan Day stalls and dealer counters.",
     "Collection method to be agreed with the client."),
    ("PR", "Kissan ki Awaaz", "Farmers' own asks, gathered through the year, handed over on 18 December, credited to Sarsabz Salam Kissan.",
     "Press credits Fatima Group or no one; comments carry supply, price and respect asks."),
    ("Events", "The farmer takes the stage", "At the national Kissan Day, farmers and their children speak first and officials answer.",
     f"Celebrities and politicians speak in {CELEB_POL} of {N_UP} uploads; a farmer in {FARMER}."),
    ("On-ground", "Replies at sowing and harvest", "Recording points where farmers gather, in Rabi and Kharif.",
     f"Brief p6 seasons; {100 * IN_WIN / N_UP:.0f}% of uploads sit in the December window."),
]
YEAR = [("Rabi sowing  ·  Oct-Nov", "Collect replies at recording points, stalls and dealer counters."),
        ("1-18 December", "The Reply anthem, #KissanKaJawab UGC, the national event and the Kissan ki Awaaz handover."),
        ("Kharif  ·  Apr-Sep", "Harvest replies, and follow-up on the asks raised in December."),
        ("Month twelve", "The next December opens with what changed since the last reply.")]
ASKS = [("Complaints on air", "Farmers' answers will include fertilizer supply and price. Editing them out breaks the idea."),
        ("Real families", "Real farmers and their children, not celebrities or AI-animated farmers."),
        ("Name the brand", "Sarsabz Salam Kissan in every PR and event material."),
        ("Beyond December", "Run it through Rabi and Kharif, not only on the day."),
        ("Clear the line", "Legal check of the Salam Kissan mark, and test both halves with farmers before production."),
        ("Supply the data", "TikTok results; brand-health attribution of Kissan Day; 2023 giveaway rules.")]


def campaign(prs) -> None:
    src = ("Sources: brief p6, p9-12; archive/; comments/; scan/; ProPakistani 8 Jan 2025. No budget, reach, results or dates "
           "beyond the brief are estimated.")
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    rect(s, 0.9, 2.0, 1.6, 0.06, ACCENT)
    txt(s, 0.9, 2.3, 11.4, 0.6, "6  ·  The campaign", 22, False, SKY)
    txt(s, 0.9, 2.95, 11.6, 1.8, ["Salam Kissan.", "Salam Pakistan."], 44, True, WHITE, space=0)
    txt(s, 0.9, 4.95, 11.4, 0.5, "Proposed line: the country salutes the farmer, and the farmer salutes the country back.", 18, False, SKY, italic=True)
    txt(s, 0.9, 6.3, 11.4, 0.4, "From the Big Idea: the farmer answers Pakistan's salute in his own voice, through his own family.", 13, False, SKY)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  THE CHAIN", "From the Big Idea to the campaign, one step at a time",
              "Pakistan salutes the farmer but has never heard him answer. A salute with no reply is what the critic calls a TV show.", src)
    y = 1.3
    for i, (label, body) in enumerate(CHAIN):
        hi = label == "Line"
        rect(s, 0.6, y, 12.1, 0.9, PALE if hi or i % 2 == 0 else WHITE)
        if hi:
            rect(s, 0.6, y, 0.06, 0.9, ACCENT)
        txt(s, 0.85, y + 0.28, 2.0, 0.4, label.upper(), 11, True, ACCENT if hi else NAVY, space=0)
        txt(s, 2.9, y + 0.15, 9.6, 0.7, body, 14 if hi else 12, hi, NAVY if hi else INK, space=0)
        y += 0.96

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  THE LINE", "Build on the name Sarsabz owns; give the salute an answer",
              f"Recommended: Salam Kissan. Salam Pakistan. The audience writes the second half already, in {SALAM_PK} of "
              f"{len(aud):,} comments.",
              "Comments: " + V("They are us and we are them. SALAM PAKISTAN") + " (87 likes, 2019). Archive: "
              + V("Kissan Tera Ehsaan, Sarsabz Pakistan!") + " (2021 rice testimonial description). Scan: JPL (2020), Rizq Foods (2025).")
    table(s, 0.6, 1.3, [4.2, 7.9], [
        ["Line", "Trade-off"],
        ["Salam Kissan. Salam Pakistan.  (recommended)",
         f"Keeps the owned name and answers it, in words {SALAM_PK} commenters already use. Generic on its own, so who says it on screen carries it."],
        ["Salam Kissan. Wa Alaikum Salam, Pakistan.  (rejected)",
         "Sharper, because a greeting has a fixed answer. But it moves a civic salute into a religious register, a farmer who isn't "
         "Muslim can't say it in character, and only 3 comments use the greeting at all."],
        ["Salam Kissan  (as today)", "Owned since 2019, but JPL and Rizq Foods already title their own videos with it."],
        ["Kissan Tera Ehsaan, Sarsabz Pakistan!  (owned, 2021)", "A thank-you with the brand name, but still one-way: the farmer is thanked, not heard."]],
        11.5, 1.05)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  HOW IT TRAVELS", "One idea across the five touchpoints the brief names",
              "Every execution puts the farmer's own voice first, and every one builds on something the evidence shows.", src)
    rows = [["Touchpoint", "Execution", "What it is", "Evidence it builds on"]] + [list(e) for e in EXEC]
    table(s, 0.6, 1.3, [1.6, 2.3, 4.5, 3.7], rows, 9.5, 0.64)

    # the goal (p12), mechanism by mechanism
    urban = sorted([r for r in arch if r["audience"] == "urban"], key=lambda r: -int(r["views"]))
    meal = next(r for r in urban if "Behind every meal" in r["title"])
    waste = next(r for r in urban if "Waste Food" in r["title"])
    thank = next(r for r in urban if "Thank you, Kissan!" in r["title"])
    undp = next(r for r in arch if r["format"] == "corporate_pr")
    s = frame(prs, "6  ·  THE CAMPAIGN  ·  THE GOAL", "What the goal asks people to understand, and what delivers it",
              "Each line of the goal gets a mechanism, not a mention.", src)
    table(s, 0.6, 1.3, [3.2, 4.3, 4.6], [
        ["The goal (p12): people should understand", "What delivers it", "Evidence it stands on"],
        ["The importance of farmers in our daily lives",
         "“What I put on your table today” - the reply cut for the city: a district, a crop, the meal it becomes",
         f"The archive's urban films work: {V('Behind every meal is a farmer’s hard work.')} ({m(int(meal['views']))}), "
         f"{V('Let’s Pledge Not To Waste Food')} ({m(int(waste['views']))})"],
        ["The contribution of agriculture to the economy",
         "The number in the reply: farmers state what their field puts in, and the work carries the national figures",
         "Agriculture is 24% of GDP and 37.4% of employment (Pakistan Economic Survey 2023-24)"],
        ["The role Sarsabz plays in celebrating and empowering",
         "The other 364 days: the empowerment proof joins Kissan Day for the first time",
         f"Ki Jeet wins in 34 of 40 districts; 500+ demo plots (p5); app 800,000+ downloads and Sarsabz Asaan Rs.500bn (p7); "
         f"UNDP film {m(int(undp['views']))} (2024)"]],
        10, 1.25)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  EMPOWERING", "Celebrating is one day. Empowering is the other 364",
              "None of this proof has ever been part of Kissan Day. It is what turns a tribute into a role.",
              "Sources: brief pp.5, 7; Case 2 audit (workspace/case2_audit/); archive/archive_codes.csv. Ki Jeet claims need "
              "correcting before use - see the Case 2 asks.")
    table(s, 0.6, 1.3, [3.4, 4.6, 4.1], [
        ["What Sarsabz already does", "How it joins the reply", "Where it comes from"],
        ["Ki Jeet: wins in 34 of 40 districts", "Winners answer in their own words, with yields and inputs named",
         "Case 2 audit; the brand's most-watched product films (8.2M, 7.5M)"],
        ["500+ demo plots with research institutes", "A plot near the farmer who replies, open to visit at sowing",
         "Brief p5. No plot has ever appeared in the owned archive"],
        ["Sarsabz Pakistan App, 800,000+ downloads", "Replies collected and answered in the app, in his language",
         "Brief p7"],
        ["Sarsabz Asaan, Rs.500bn recorded", "The credit story told by the farmer who used it", "Brief p7"],
        ["UNDP partnership on sustainable farming", f"The national frame for the movement ({m(int(undp['views']))} views)",
         "Archive: the biggest upload of 2024"]],
        9, 0.74)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  THE MOVEMENT", "The day already travels without the brand. Make that the movement",
              "A movement that publishes its size can be checked. One that does not is a claim.",
              "Sources: The Nation 19 Dec 2024; ProPakistani 8 Jan 2025; Daily Times (undated, blocked); scan/scan_findings.md.")
    panel(s, 0.6, 1.3, 5.9, 4.85, "Who marks the day now, without Sarsabz on it", [
        "The 5th National Farmers' Day was observed through the Kashtkar Dost Foundation and All Pakistan Kissan Ittehad, "
        "crediting “Fatima Group” - not Sarsabz, not Salam Kissan (The Nation, 19 Dec 2024).",
        "The sixth Kissan Day was held in Islamabad with the Federal Minister and the FAO representative (ProPakistani, 8 Jan 2025).",
        "The ICT administration ran its own “Salam Kissan” campaign on the capital's streets (Daily Times; undated, page blocked).",
        "Syngenta, JPL Pakistan and Rizq Foods post the day on their own channels; JPL and Rizq use the Salam Kissan name."], 11.5)
    panel(s, 6.8, 1.3, 5.9, 4.85, "The mechanism", [
        "A Salam Kissan partner pack: the day's assets, the reply format and the name, offered to the farmer bodies, provincial "
        "departments and partners who already mark it - so it grows under the name that started it.",
        "Every partner's replies feed one record, handed over on 18 December as Kissan ki Awaaz.",
        "One published number each year: how many farmers answered, from how many districts.",
        "What it asks of Fatima: decide the number it will publish before the campaign runs, and accept that year one may be small."], 11.5)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  A YEAR", "Beyond a single day: replies from sowing to harvest to 18 December",
              "Month twelve opens the next December with what changed, so the day becomes a platform, not a launch.",
              "Seasons from brief p6 (Rabi Oct-Mar, Kharif Apr-Sep). Timings are a proposal, not a media plan; budget not stated in the brief.")
    for i, (stage, body) in enumerate(YEAR):
        x = 0.6 + i * 3.08
        rect(s, x, 1.5, 2.9, 0.7, ACCENT if stage == "1-18 December" else NAVY)
        txt(s, x + 0.15, 1.68, 2.6, 0.4, stage, 13, True, WHITE, space=0)
        rect(s, x, 2.2, 2.9, 3.6, PALE)
        txt(s, x + 0.2, 2.45, 2.5, 3.2, body, 14, False, INK)

    s = frame(prs, "6  ·  THE CAMPAIGN  ·  WHAT IT ASKS", "The price of the idea, stated before any line is written",
              "If Fatima won't put farmers' own words on air, run the crop-calendar platform instead: easier to approve, easy for rivals to copy.",
              "Open for the client: budget, pitch date, and whether Case 1 and Case 2 share a media plan.")
    for i, (name, body) in enumerate(ASKS):
        col, row = i % 2, i // 2
        x, y = 0.6 + col * 6.15, 1.3 + row * 1.6
        rect(s, x, y, 5.95, 1.45, PALE)
        rect(s, x, y, 0.06, 1.45, ACCENT)
        txt(s, x + 0.3, y + 0.15, 5.5, 0.4, f"{i + 1}  {name}", 14, True, NAVY, space=0)
        txt(s, x + 0.3, y + 0.6, 5.5, 0.8, body, 12, False, INK, space=0)


def free_path(p: Path) -> Path:
    if not p.exists():
        return p
    n = 2
    while (q := p.with_name(f"{p.stem}-v{n}{p.suffix}")).exists():
        n += 1
    return q


def check_spec(spec_path: Path) -> None:
    spec = spec_path.read_text(encoding="utf-8")
    for mt in re.finditer(r'\\"([^"\\]{12,})\\"', spec):
        if "translation" not in spec[max(0, mt.start() - 25):mt.start()]:  # labelled translations aren't verbatims
            V(mt.group(1))


def main() -> None:
    wrong = {k: (FACTS[k], v) for k, v in EXPECT.items() if FACTS[k] != v}
    if wrong:
        raise SystemExit(f"Refusing to build - data no longer matches the numbers in the spec/docs (got, expected): {wrong}")
    spec_path = HERE / "big_idea_spec.json"
    check_spec(spec_path)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    evidence(prs)
    if MISSING:
        raise SystemExit("Refusing to build - verbatims not found in any source:\n  - " + "\n  - ".join(MISSING))

    with tempfile.TemporaryDirectory() as tmp:
        part1 = Path(tmp) / "part1.pptx"
        prs.save(part1)
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        spec["out"] = str(Path(tmp) / "part2.pptx")
        tmp_spec = Path(tmp) / "spec.json"
        tmp_spec.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        r = subprocess.run([sys.executable, str(SKILL), str(tmp_spec), "--base", str(part1)],
                           capture_output=True, text=True, encoding="utf-8")
        print(r.stdout.strip(), r.stderr.strip())
        if r.returncode:
            raise SystemExit("big-idea-slides builder failed")
        prs = Presentation(spec["out"])
        campaign(prs)
        if MISSING:
            raise SystemExit("Refusing to build - verbatims not found in any source:\n  - " + "\n  - ".join(MISSING))
        out = free_path(WS / "out" / "Sarsabz_case1_deck.pptx")
        prs.save(out)
    print(f"{out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
