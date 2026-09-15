"""Build the brand-by-brand category audit deck.

    uv run --with python-pptx --with pillow python workspace/audit/build_deck.py

Date-stamped, never overwritten - a same-day rebuild becomes -v2, -v3.
"""

from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

AUDIT = Path(__file__).parent
OUT = AUDIT.parent / "out"

NAVY = RGBColor(0x10, 0x2A, 0x43)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x74, 0x80)
RULE = RGBColor(0xD8, 0xDD, 0xE3)
ACCENT = RGBColor(0xC8, 0x10, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

RING_LABEL = {"client": "CLIENT", "direct": "DIRECT COMPETITOR",
              "adjacent": "ADJACENT / TAKAFUL", "substitute": "SUBSTITUTE"}

# brand -> (headline read, [what they are saying], the strategic note)
READ = {
    "State Life Insurance Corporation of Pakistan": (
        "A real campaign, buried in institutional PR.",
        ["PLAN AHEAD WITH STATE LIFE - three executions: education before the admission "
         "form, healthcare before you need care, opportunity when you're prepared.",
         "Footer already claims 'Pakistan's Largest Life and Health Insurer'.",
         "E-Kachehri: 745 complaints received, 744 resolved.",
         "The other nine posts: two plaque handovers, a conference panel, a flag-raising, "
         "an Eid greeting, a fellowship visit, a product-range chart."],
        "The channel reads as a ministry newsletter with three advertisements inserted. "
        "The strongest asset in the grid - 744 of 745 - is filed as paperwork."),
    "EFU Life Assurance": (
        "Has left the category it leads.",
        ["Seven of twelve posts are loyalty offers: 25% off a burger, 30% off diagnostics, "
         "20% off a clinic, 1000 EFU Coins, Rs.500 KraveMart vouchers.",
         "Thrive and WIN apps carry the brand; insurance is barely mentioned.",
         "Historic work was duty and dignity - Zaroori Hai, Izzat say Nazrain Milain.",
         "One CEO post, 'Standing Strong Together', is the only trust claim."],
        "The largest private life insurer now advertises like a loyalty card. The duty "
        "territory it built over nineteen years is vacated."),
    "Jubilee Life Insurance": (
        "Owns the proof territory - and then abandoned it.",
        ["PKR 54 BILLION CLAIMS PAID IN 2025 / 'Behind every claim is a promise fulfilled'.",
         "Pinned to the top of the grid. Dated June.",
         "Everything after it: 40% off a laboratory, 10% off radiology, 40% off at a cafe, "
         "and reposted fitness-influencer reels.",
         "Long-running platform: 'A Samajhdar Faisla for a Befiker Mustaqbil'."],
        "The only competitor standing in the proof territory is standing there with one "
        "June post. Claimed, not held."),
    "Adamjee Life Assurance": (
        "Heritage and convenience, no proposition.",
        ["'The Home to Every Life' - map of Pakistan, heritage illustration.",
         "70 years and a Decade of Dedication award.",
         "'Kabhi bhi Kahin bhi' - Quick Pay app, 'branch visit bhool jao'.",
         "Three film stills with no readable message; four event photos."],
        "Ten of twelve posts carry no call to action. Anniversary and app convenience "
        "are doing the work a proposition should do."),
    "IGI Life Insurance": (
        "The best objection work in the category - aimed at the wrong objection.",
        ["MYTH series: 'I only need life insurance when I have a family.'",
         "MYTH series: 'My child's education is still years away.'",
         "'WHO ARE YOU PROTECTING TODAY?'",
         "Vitality rewards: PKR 500 weekly with foodpanda."],
        "Three of twelve ads handle objections directly - the highest rate in the set. "
        "But they tackle purchase timing, never whether the insurer pays."),
    "Askari Life Assurance": (
        "The sharpest platform in the category.",
        ["IRAADA KARO / Jeeto Har Ghari - Urdu-first, bold typography, intention as the "
         "organising idea.",
         "'Ek chhotay se iraade se shuru hota hai'.",
         "Chat UI: 'You added peace of mind / Uncertainty: Removed from chat'.",
         "'PLAN AHEAD WITH ASKARI LIFE EDUCATION PLAN'."],
        "623 followers, and the most coherent creative platform in the sample. Note the "
        "collision: 'Plan ahead' is the same phrase State Life is using."),
    "TPL Life Insurance": (
        "Narrow, affluent, and uncontested.",
        ["Globewell: access to medical providers in 185+ countries.",
         "Coverage of up to $500,000; 24/7 worldwide medical assistance.",
         "'From Dubai to London, Globewell has got your back.'",
         "Sehat Zindagi for the mass market: PKR 1,250,000, nationwide."],
        "Six of twelve ads claim protection - the highest in the set. Owns international "
        "healthcare cleanly. No proof device anywhere."),
    "Pak-Qatar Family Takaful": (
        "Borrowing the association State Life owns outright.",
        ["Appointed by the Federal Government as Authorized Pension Fund Manager.",
         "'RETIREMENT DESERVES CERTAINTY - Lifetime Kafalat Plan. A guaranteed halal "
         "pension for life.'",
         "'THE SECOND PAYCHEQUE' - Mahana Bachat & Takaful Flexi.",
         "AA / AM2+ ratings badge on every single creative."],
        "The only brand using a proof device systematically - 11 of 12 ads carry the "
        "rating. And it advertises federal appointment as a credential."),
    "Dawood Family Takaful": (
        "The sharpest objection-handling ad in the whole sample.",
        ["'TIE YOUR CAMEL AND TRUST IN ALLAH' - Prophet Muhammad (PBUH), Jami at "
         "Tirmidhi 2517.",
         "'Takaful or Insurance? CHOOSE THE HALAL WAY!'",
         "'TAKAFUL COVERS MORE THAN YOU THINK - full-circle protection.'",
         "Heavy patriotic calendar content: Youm-e-Takbeer, air force, Arshad Nadeem."],
        "Answers the fatalism objection with scripture. A 545-follower Takaful operator "
        "is doing braver category work than any of the big four."),
    "EFU Life Window Takaful": (
        "The state association, borrowed again.",
        ["'STRENGTHENING PAKISTAN'S RETIREMENT FUTURE WITH GOVERNMENT INITIATIVES' - "
         "Federal Hemayah Pension Fund.",
         "Hemayah x EFU General x BankIslami - micro and nano Takaful.",
         "mHealth digital healthcare platform.",
         "'Born from a DREAM' heritage post."],
        "Five of twelve posts are video with no loaded poster frame. Like Pak-Qatar, "
        "reaches for federal association as a credibility device."),
    "Al Meezan Investments": (
        "The loudest voice in the set - and not an insurer.",
        ["81.8K followers - four times EFU, thirty times State Life.",
         "Genuine analysis: KSE100 P/E re-rating charts, Market Outlook 2026.",
         "'HOME BUILDER PLAN - some doors open long before you reach them.'",
         "'MYTH & FACT' with an Al Meezan Buddy mascot."],
        "A Shariah-compliant fund manager out-publishes the entire life category and "
        "competes for the same protection-and-savings wallet."),
    "UBL Fund Managers": (
        "Liquidity as the counter-offer.",
        ["'Fauri Paisa - your money, when you need it. Instant redemption up to "
         "PKR 1,000,000.'",
         "'Faisla yaqeen ka' campaign line.",
         "Invest via RAAST; Al-Ameen rebrand.",
         "Three of twelve posts are recruitment ads."],
        "21.6K followers. Attacks life insurance where it is weakest - money locked away "
        "for decades - by selling instant access."),
}


def txbox(slide, l, t, w, h, text, size=12, bold=False, color=INK,
          align=PP_ALIGN.LEFT, space=6, font="Segoe UI"):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        p.space_after = Pt(space)
        f = p.runs[0].font
        f.size, f.bold, f.color.rgb, f.name = Pt(size), bold, color, font
    return tb


def rect(slide, l, t, w, h, fill):
    from pptx.enum.shapes import MSO_SHAPE
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    with (AUDIT / "profiles.csv").open(encoding="utf-8", newline="") as fh:
        profiles = {r["brand"]: r for r in csv.DictReader(fh)}
    with (AUDIT / "handles.csv").open(encoding="utf-8-sig", newline="") as fh:
        order = list(csv.DictReader(fh))

    wb = openpyxl.load_workbook(AUDIT / "capture.xlsx")
    ws = wb["Ads"]
    hdr = [c.value for c in ws[1]]
    rows = [dict(zip(hdr, [c.value for c in r])) for r in ws.iter_rows(min_row=2) if r[1].value]

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    W = prs.slide_width

    # ---- title ----
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, prs.slide_height, NAVY)
    txbox(s, Inches(0.9), Inches(2.2), Inches(11), Inches(1),
          "How Pakistan's life insurance category talks", 40, True, WHITE)
    txbox(s, Inches(0.9), Inches(3.3), Inches(11), Inches(1),
          "A brand-by-brand communications audit", 20, False, RGBColor(0x9F, 0xB4, 0xC7))
    txbox(s, Inches(0.9), Inches(5.6), Inches(11), Inches(1),
          ["144 ads  ·  12 brands  ·  collected 27 August 2026",
           "State Life Insurance Corporation of Pakistan  ·  EPADS P78118"],
          13, False, RGBColor(0x9F, 0xB4, 0xC7))

    # ---- sample frame ----
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, Inches(1.0), NAVY)
    txbox(s, Inches(0.6), Inches(0.3), Inches(10), Inches(0.5),
          "What we looked at", 24, True, WHITE)
    txbox(s, Inches(0.6), Inches(1.4), Inches(5.9), Inches(4),
          ["The twelve most recent posts from each brand's public Instagram, collected on "
           "27 August 2026.",
           "Every artwork was downloaded and coded from the image, not from the caption.",
           "119 of 144 coded from the creative. 25 left unclear where the message could not "
           "be read - none were guessed.",
           "Postal Life and the Central Directorate of National Savings have no Instagram "
           "presence at all. Recorded as absences, not gaps."], 13, False, INK, space=10)
    rect(s, Inches(6.9), Inches(1.4), Inches(5.8), Inches(0.05), ACCENT)
    txbox(s, Inches(6.9), Inches(1.7), Inches(5.8), Inches(3),
          ["A limit worth stating first",
           "Logged-out access caps collection at twelve posts per brand. Every brand is "
           "therefore 8% of the sample by construction.",
           "That means share of voice cannot be measured here, and no share-of-voice "
           "number appears in this deck.",
           "Everything shown is organic owned content. No paid creative, no spend, no reach."],
          13, False, INK, space=10)
    s.shapes[-1].text_frame.paragraphs[0].runs[0].font.bold = True

    # ---- how the category talks ----
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, Inches(1.0), NAVY)
    txbox(s, Inches(0.6), Inches(0.3), Inches(10), Inches(0.5),
          "How the category talks", 24, True, WHITE)
    stats = [("78%", "of ads use no proof device at all", "112 of 144. No statistic, rating, testimonial or guarantee."),
             ("2", "ads cite a claims number", "Jubilee's PKR 54 billion. State Life's 744 of 745."),
             ("6%", "address the core objection", "9 of 144. The sharpest is from a 545-follower Takaful operator."),
             ("88", "ads carry no call to action", "Of 144.")]
    x = Inches(0.6)
    for big, mid, small in stats:
        rect(s, x, Inches(1.5), Inches(2.95), Inches(0.05), ACCENT)
        txbox(s, x, Inches(1.8), Inches(2.95), Inches(0.9), big, 40, True, NAVY)
        txbox(s, x, Inches(2.75), Inches(2.95), Inches(0.7), mid, 13, True, INK)
        txbox(s, x, Inches(3.5), Inches(2.95), Inches(1.2), small, 11, False, MUTED)
        x += Inches(3.12)
    txbox(s, Inches(0.6), Inches(5.0), Inches(12.1), Inches(2),
          ["The category has stopped selling insurance and started selling discounts.",
           "Three of the four largest private life insurers - EFU, Jubilee and IGI - are "
           "running the same wellness-rewards mechanic. Seven of EFU's twelve recent posts "
           "are loyalty offers. Patriotism is the second-most-crowded territory and is "
           "almost entirely decorative: of 17 national-trust claims, 13 sit in the pride "
           "register - flags, cakes, a fighter jet - and only four give a reason to buy."],
          14, False, INK, space=8)
    s.shapes[-1].text_frame.paragraphs[0].runs[0].font.bold = True
    s.shapes[-1].text_frame.paragraphs[0].runs[0].font.size = Pt(17)

    # ---- one slide per brand ----
    for h in order:
        brand = h["brand"]
        if brand not in READ:
            continue
        prof = profiles.get(brand, {})
        rs = [r for r in rows if r["brand"] == brand]
        headline, saying, note = READ[brand]

        s = prs.slides.add_slide(blank)
        rect(s, 0, 0, W, Inches(1.15), NAVY)
        txbox(s, Inches(0.55), Inches(0.22), Inches(8.6), Inches(0.5), brand, 22, True, WHITE)
        ring = RING_LABEL.get(h.get("ring", ""), h.get("ring", "").upper())
        followers = prof.get("followers") or "-"
        txbox(s, Inches(0.55), Inches(0.72), Inches(8.6), Inches(0.35),
              f"@{h['handle']}   ·   {ring}   ·   {followers} followers   ·   {len(rs)} ads coded",
              12, False, RGBColor(0x9F, 0xB4, 0xC7))

        sheet = AUDIT / "contact" / f"{h['handle']}.jpg"
        if sheet.exists():
            s.shapes.add_picture(str(sheet), Inches(0.55), Inches(1.5), width=Inches(6.7))

        left = Inches(7.6)
        wid = Inches(5.2)
        rect(s, left, Inches(1.5), wid, Inches(0.05), ACCENT)
        txbox(s, left, Inches(1.75), wid, Inches(0.8), headline, 17, True, NAVY)
        txbox(s, left, Inches(2.75), wid, Inches(2.6),
              [f"·  {x}" for x in saying], 11.5, False, INK, space=7)

        rect(s, left, Inches(5.45), wid, Inches(1.45), RGBColor(0xF2, 0xF5, 0xF8))
        txbox(s, left + Emu(120000), Inches(5.6), wid - Emu(240000), Inches(1.2),
              note, 12, False, NAVY, space=4)

        # counted facts, so the read is anchored
        nocta = sum(1 for r in rs if r["cta_channel"] == "none")
        noproof = sum(1 for r in rs if r["proof_device"] == "none")
        obj = sum(1 for r in rs if r["addresses_objection"] == "yes")
        txbox(s, Inches(0.55), Inches(6.95), Inches(6.7), Inches(0.4),
              f"{noproof}/{len(rs)} ads with no proof device   ·   {obj}/{len(rs)} address the "
              f"core objection   ·   {nocta}/{len(rs)} with no call to action",
              10.5, False, MUTED)

    # ---- the finding ----
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, prs.slide_height, NAVY)
    rect(s, Inches(0.9), Inches(1.5), Inches(1.6), Inches(0.06), ACCENT)
    txbox(s, Inches(0.9), Inches(1.9), Inches(11.4), Inches(1.4),
          "The vacant territory is proof.", 40, True, WHITE)
    txbox(s, Inches(0.9), Inches(3.2), Inches(11.4), Inches(2.6),
          ["Not trust as a feeling. Trust as an audited, publishable, repeatable number.",
           "78% of category advertising offers no proof of anything. Two ads in 144 cite a "
           "claims figure. The question every buyer actually asks - will they pay? - is "
           "answered in passing, by one competitor, once, in June.",
           "State Life already owns the strongest number in the category and files it as "
           "paperwork."],
          15, False, RGBColor(0xC7, 0xD4, 0xE0), space=12)

    # ---- what State Life can own ----
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, Inches(1.0), NAVY)
    txbox(s, Inches(0.6), Inches(0.3), Inches(10), Inches(0.5),
          "What State Life can credibly own", 24, True, WHITE)
    rect(s, Inches(0.6), Inches(1.4), Inches(5.6), Inches(2.2), RGBColor(0xF2, 0xF5, 0xF8))
    txbox(s, Inches(0.9), Inches(1.75), Inches(5.0), Inches(1.6),
          ["744 of 745", "complaints resolved.",
           "Published in a grievance-redress announcement, in a government template, "
           "to 2,499 followers."], 15, False, NAVY, space=8)
    p0 = s.shapes[-1].text_frame.paragraphs[0].runs[0].font
    p0.size, p0.bold, p0.color.rgb = Pt(38), True, ACCENT
    txbox(s, Inches(6.6), Inches(1.4), Inches(6.1), Inches(2.4),
          ["Stronger than Jubilee's number, not weaker.",
           "Rupees paid measures volume. Disputes resolved measures behaviour when the "
           "customer pushes back. That is the objection."], 14, False, INK, space=10)
    s.shapes[-1].text_frame.paragraphs[0].runs[0].font.bold = True

    txbox(s, Inches(0.6), Inches(4.0), Inches(12.1), Inches(0.4),
          "Three things no private competitor can copy", 15, True, NAVY)
    items = [("Sovereign backing",
              "A state-owned insurer's promise is underwritten by the state. Pak-Qatar and "
              "EFU both advertise borrowed federal association - the thing State Life owns."),
             ("A statutory grievance obligation",
              "E-Kachehri is a duty, not a campaign. It generates the number every quarter "
              "whether or not marketing wants it."),
             ("Scale and heritage",
              "'Pakistan's Largest Life and Health Insurer' is already in the footer of the "
              "existing work, doing nothing.")]
    x = Inches(0.6)
    for t, b in items:
        rect(s, x, Inches(4.5), Inches(3.9), Inches(0.05), ACCENT)
        txbox(s, x, Inches(4.75), Inches(3.9), Inches(0.5), t, 13, True, NAVY)
        txbox(s, x, Inches(5.35), Inches(3.9), Inches(1.6), b, 11.5, False, INK)
        x += Inches(4.08)
    txbox(s, Inches(0.6), Inches(6.85), Inches(12.1), Inches(0.4),
          "The one competitor standing in this territory is standing there with a single "
          "June post. This is not a contested position. It is an unattended one.",
          12, True, ACCENT)

    stem = f"SLIC_P78118_category_audit_{date.today().isoformat()}"
    path = OUT / f"{stem}.pptx"
    n = 2
    while path.exists():
        path = OUT / f"{stem}-v{n}.pptx"
        n += 1
    prs.save(path)
    print(f"{path}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")


if __name__ == "__main__":
    main()
