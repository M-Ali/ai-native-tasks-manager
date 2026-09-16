"""The point of view: the deck that answers "What We Expect from the Pitch" (brief p14).

    uv run --no-project --with python-pptx python workspace/pov/build_pov_deck.py

Writes workspace/out/Sarsabz_point_of_view.pptx (never overwrites: a clash becomes -v2).

Neither case deck answers p14. Case 1 is Salam Kissan, Case 2 is the audit and the tactical campaign;
the seven expectations sit above both - portfolio, Bubber Sher, category dynamics, the emotional /
functional balance, measurement and long-term delivery. This builds that third deck.

THE GUARD: every figure on a slide comes from FACTS, and every FACTS entry carries its source.
fact("agri_gdp") returns the number and registers the source for the slide's footer, so a number
cannot reach a slide without one. External figures are quoted as the source states them, and where
two sources disagree (PBS says ~24% of GDP and half the labour force; the Economic Survey says 23.4%
and 33.1%) the disagreement goes on the slide rather than being averaged away.
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = Path(__file__).parent
WS = HERE.parent

NAVY = RGBColor(0x10, 0x2A, 0x43)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x6B, 0x74, 0x80)
ACCENT = RGBColor(0xC8, 0x10, 0x2E)
PALE = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SKY = RGBColor(0x9F, 0xB4, 0xC7)

PES = "Pakistan Economic Survey 2025-26, ch.2 Agriculture (finance.gov.pk), retrieved 16 Sep 2026"
PBS = "Pakistan Bureau of Statistics, 'Agriculture Sector of Pakistan' (pbs.gov.pk), undated page"
PACRA_F = "PACRA rating report, Fatima Fertilizer Company Ltd, 24 Jul 2026"
PACRA_FF = "PACRA press release, Fatimafert Ltd, 22 Apr 2020"
PACRA_FFC = "PACRA rating report, Fauji Fertilizer Company Ltd, 31 Jul 2026"
FATIMA = "fatima-group.com, brand and capacity pages, retrieved 16 Sep 2026"
BRIEF = "Fatima Creative Marketing Brief 2026"
OURS = "Our own analysis: 200 coded competitor uploads, 3,005 farmer comments, 76-upload Salam Kissan archive"

# value -> (text, source). Nothing numeric reaches a slide except through fact().
FACTS: dict[str, tuple[str, str]] = {
    "agri_gdp": ("23.4% of GDP", PES),
    "agri_emp": ("33.1% of employment", PES),
    "agri_growth": ("sector growth 2.89% in 2025-26, crops 1.44%", PES),
    "pbs_gdp": ("about 24% of GDP and 'half of employed labour force'", PBS),
    "offtake": ("3,795 thousand nutrient tonnes, +11.4% year on year (Jul-Mar FY2026)", PES),
    "n_p_k": ("nitrogen 3,035k tonnes (+14.8%), phosphate 712k (-1.9%), potash 48k", PES),
    "phos_price": ("phosphate offtake fell on high prices", PES),
    "bag_burden": ("Rs 100 more per 50kg bag = Rs 20 billion more on farmers", PES),
    "kissan_card": ("Punjab's Kissan Card improved cotton fertilizer application", PES),
    "tractors": ("limited tractor and machinery availability remains a key constraint", PES),
    "ffc_share": ("FFC urea share 56% and DAP 66% in 1HCY26 (43% and 62% in CY25)", PACRA_FFC),
    "industry_offtake": ("industry offtake around 9.3mln MT in CY25", PACRA_FFC),
    "ffc_capacity": ("FFC capacity 2,599 KT urea and 650 KT DAP, utilisation above 124%", PACRA_FFC),
    "oligopoly": ("'a prominent player in Pakistan's oligopolistic fertilizer industry'", PACRA_F),
    "portfolio": ("'Sarsabz Urea, Sarsabz DAP, Sarsabz Nitrophos, Sarsabz Calcium Ammonium Nitrate, "
                  "Bubbersher Urea and Bubbersher DAP'", PACRA_F),
    "flagships": ("'widely recognized for its flagship brands, Sarsabz and Bubbersher'", PACRA_F),
    "carveout": ("the Multan plant was carved out to Pakarab Fertilizers Ltd with effect from 1 Jan 2025", PACRA_F),
    "demand_risk": ("demand 'exposed to fluctuations in farm economics, crop prices, government subsidy "
                    "policies, and volatility in global commodity and gas prices'", PACRA_F),
    "bs_plant": ("Bubber Sher urea is made at FatimaFert, Sheikhupura; the plant's nameplate capacity was "
                 "445,500 MT", PACRA_FF),
    "bs_oldest": ("'one of the oldest in Pakistan and widely recognized and trusted by farmers'; the brand was "
                  "acquired by buying Bubber Sher (Private) Ltd; its DAP is imported", FATIMA),
    "bs_zone": ("the brief calls it 'a regional brand... in North Zone' (p20); the company's own site describes it "
                "marketed across Pakistan with no regional limit", BRIEF + "; " + FATIMA),
    "group_capacity": ("group capacity: urea 1,045,000 MT, CAN 920,000 MT, NP 820,000 MT", FATIMA),
    "engro_vs": ("Engro answers a farmer objection in 58% of its 110 uploads; Sarsabz in 6% of its 90", OURS),
    "no_proof": ("no brand in 200 coded uploads shows a result on a real field beside the old way", OURS),
    "mixing": ("mixing and compatibility is the largest subject among farmers: 220 of 1,433 comments", OURS),
    "farmer_voice": ("a farmer speaks in 9 of 76 Salam Kissan uploads; celebrities and politicians in 30", OURS),
    "no_channel": ("Bubber Sher has no official YouTube channel; the 'Bubber Sher' Facebook page is a rock band", OURS),
}
USED: list[str] = []


def fact(key: str) -> str:
    USED.append(key)
    return FACTS[key][0]


def sources_line(keys: list[str]) -> str:
    seen = []
    for k in keys:
        # a fact can cite two documents ("brief; site"); split before de-duplicating, or the
        # footer prints the same source twice
        for s in FACTS[k][1].split("; "):
            if s not in seen:
                seen.append(s)
    return "Sources: " + "  ·  ".join(seen)


def cap(s: str) -> str:
    """Facts are written to sit mid-sentence; capitalise when one opens a bullet."""
    return s[:1].upper() + s[1:] if s else s


def txt(slide, l, t, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, space=6, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(line) or " "  # an empty cell yields no run, and styling it would crash
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
        txt(s, 0.6, 7.05, 12.1, 0.38, source, 8, False, MUTED, space=0)
    return s


def panel(s, l, t, w, h, title, lines, size=11):
    rect(s, l, t, w, h, PALE)
    rect(s, l, t, 0.06, h, ACCENT)
    txt(s, l + 0.25, t + 0.15, w - 0.45, 0.32, title.upper(), 10, True, ACCENT, space=0)
    txt(s, l + 0.25, t + 0.55, w - 0.45, h - 0.72, lines, size, False, INK, space=7)


def table(s, l, t, widths, rows, size=10.5, row_h=0.62, head_h=0.45):
    y = t
    for i, row in enumerate(rows):
        hdr = i == 0
        h = head_h if hdr else row_h
        if hdr:
            rect(s, l, y, sum(widths), h, NAVY)
        elif i % 2 == 0:
            rect(s, l, y, sum(widths), h, PALE)
        x = l
        for w, cell in zip(widths, row):
            txt(s, x + 0.12, y + 0.1, w - 0.24, h - 0.15, cell, size, hdr, WHITE if hdr else INK, space=0)
            x += w
        y += h
    if y > 6.3:
        raise SystemExit(f"table runs to {y:.2f} in and would cover the conclusion bar")
    return y


def build(prs) -> None:
    # 1 cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    rect(s, 0.9, 1.7, 1.6, 0.06, ACCENT)
    txt(s, 0.9, 1.95, 11.4, 0.5, "Fatima Fertilizer creative pitch", 18, False, SKY)
    txt(s, 0.9, 2.5, 11.6, 1.0, "Our point of view", 48, True, WHITE)
    txt(s, 0.9, 3.6, 11.4, 0.6, "Where the brands go, and how communication earns the farmer's relationship", 22, False, WHITE, italic=True)
    txt(s, 0.9, 4.8, 11.4, 1.2, ["1  The portfolio   ·   2  Bubber Sher   ·   3  The category in numbers",
                                 "4  The audience   ·   5  What communication can do   ·   6  Emotion x proof",
                                 "7  How we would measure it   ·   8  The long term"], 14, False, SKY)
    txt(s, 0.9, 6.6, 11.6, 0.4, "Answering 'What We Expect from the Pitch' (brief p14)  ·  internal working draft, 16 September 2026",
        11, False, SKY)

    # 2 portfolio
    k = ["portfolio", "flagships", "carveout", "group_capacity", "no_channel"]
    s = frame(prs, "1  ·  THE PORTFOLIO", "Where each brand goes, and why",
              "One house, three jobs: Sarsabz earns the premium, Bubber Sher holds the conventional base, Pakarab CAN "
              "supplies the specialist.", sources_line(k))
    txt(s, 0.6, 1.2, 12.1, 0.4, f"Fatima's own record: {fact('portfolio')}. {fact('flagships')}.", 11, False, MUTED, space=0)
    table(s, 0.6, 1.75, [2.2, 3.3, 3.3, 3.3], [
        ["Brand", "What it is today", "The job we would give it", "What that asks"],
        ["Sarsabz", "The premium, differentiated range: NP and CAN nobody else makes at scale",
         "Carry the proof: the 10% promise shown on real fields, and the farmer's own voice on Kissan Day",
         "Publish plot results; put farmers on screen"],
        ["Bubber Sher", "Conventional urea and DAP; " + fact("bs_oldest").split(";")[0],
         "Hold the base on trust and availability, not on emotional advertising",
         "Settle the zone question; give it a presence it does not have"],
        ["Pakarab CAN", "CAN from the carved-out Multan plant: " + fact("carveout"),
         "The specialist nitrogen source, sold on agronomy to the farmers who already know CAN",
         "Decide whether it competes with Sarsabz CAN or complements it"]],
        10, 1.35)

    # 3 Bubber Sher
    k = ["bs_zone", "bs_oldest", "bs_plant", "no_channel"]
    s = frame(prs, "2  ·  BUBBER SHER", "A heritage brand with no voice, and a definition that does not agree",
              "Before a vision can be written, Fatima has to say which Bubber Sher this is: a North Zone tactical brand, "
              "or one of Pakistan's oldest national names.", sources_line(k))
    panel(s, 0.6, 1.25, 5.9, 2.4, "The contradiction we found", [
        "The brief: 'A regional brand that markets our conventional fertilizer Urea and DAP in North Zone' (p20).",
        "The company's own site: one of the oldest brands in Pakistan, 'widely recognized and trusted by farmers', "
        "described with no regional limit.",
        "Both cannot brief the same campaign. The answer changes the budget, the media map and the idea."], 11.5)
    panel(s, 0.6, 3.8, 5.9, 2.35, "What is on the record", [
        fact("bs_plant") + ".",
        fact("bs_oldest") + ".",
        fact("no_channel") + "."], 11.5)
    panel(s, 6.8, 1.25, 5.9, 4.9, "The vision we would take, if it is the national heritage brand", [
        "Bubber Sher does not need emotion; it needs presence where the conventional bag is bought. Its equity is age and "
        "trust, and it is being spent down in silence.",
        "Role: the dependable bag - available, genuine, priced as promised. Sarsabz argues yield; Bubber Sher argues certainty.",
        "Communication: dealer-first. Stickered authenticity, a visible price promise, and one campaign a year that says the "
        "brand is still made here, at Sheikhupura.",
        "The counterfeit angle is real and unowned: farmers raise fake product unprompted in the comments, and no brand answers it.",
        "If it is the North Zone tactical brand instead, the same assets shrink to a regional trade programme - and we would say "
        "so rather than inflate it."], 11)

    # 4 the category
    k = ["offtake", "n_p_k", "phos_price", "bag_burden", "ffc_share", "industry_offtake", "ffc_capacity", "oligopoly", "demand_risk"]
    s = frame(prs, "3  ·  THE CATEGORY IN NUMBERS", "A price-led, supply-led category where scale is already decided",
              "Volume is settled by capacity and price. What is not settled is who the farmer believes - and that is where "
              "communication still moves the market.", sources_line(k))
    table(s, 0.6, 1.25, [4.3, 4.2, 3.6], [
        ["What the numbers say", "The figure", "What it means for us"],
        ["The category is growing on nitrogen, shrinking on phosphate", fact("offtake") + "; " + fact("n_p_k"),
         "DAP demand follows price, not persuasion; NP and CAN are where argument still works"],
        ["Price is the farmer's first filter", fact("bag_burden") + "; " + fact("phos_price"),
         "A yield claim has to be worth more than the bag's premium, in his arithmetic"],
        ["Scale belongs to FFC", fact("ffc_share") + "; " + fact("ffc_capacity"),
         "Fatima will not win on share of volume; it can win on share of proof"],
        ["The structure is an oligopoly exposed to policy", fact("oligopoly") + "; " + fact("demand_risk"),
         "Brand preference is the only lever the company fully controls"]],
        10, 1.15)

    # 5 the audience
    k = ["agri_gdp", "agri_emp", "agri_growth", "pbs_gdp", "kissan_card", "tractors", "mixing", "farmer_voice"]
    s = frame(prs, "4  ·  THE AUDIENCE", "The sector the country depends on, and the farmer nobody answers",
              "The macro case for the farmer is settled. The relationship is not.", sources_line(k))
    panel(s, 0.6, 1.25, 6.0, 2.3, "What agriculture is worth", [
        f"Economic Survey 2025-26: {fact('agri_gdp')}, {fact('agri_emp')}, {fact('agri_growth')}.",
        f"The Bureau of Statistics page says {fact('pbs_gdp')} - it is undated, and it disagrees with the Survey. "
        "We use the Survey and show both rather than pick the flattering one."], 11.5)
    panel(s, 0.6, 3.7, 6.0, 2.45, "What is changing on the ground", [
        fact("kissan_card") + " - the state is now an input channel, and a competitor for credit for the farmer's success.",
        cap(fact("tractors")) + ".",
        "Floods in 2025 shaped the crop year, and they shaped what farmers wrote to the brand."], 11.5)
    panel(s, 6.9, 1.25, 5.8, 4.9, "What the farmer actually asks (our own reading)", [
        cap(fact("mixing")) + " - the barrier is waste, not disbelief.",
        "He asks strangers in comment threads what the maker should be telling him, and suspects the answers are paid for.",
        cap(fact("farmer_voice")) + ".",
        "He wants a per-acre answer for his soil and district, written proof, and someone who replies.",
        "Corpus limits: self-selected commenters, no age or location split - this is what people say, not a market sample."], 11)

    # 6 what communication can do
    k = ["engro_vs", "no_proof", "farmer_voice"]
    s = frame(prs, "5  ·  WHAT COMMUNICATION CAN DO", "Three jobs communication can do that price cannot",
              "Relationships are built where the farmer already argues: in the thread, at the counter, on his own field.",
              sources_line(k) + "  ·  " + BRIEF)
    for i, (head, body) in enumerate([
        ("Answer the question he asked", "The category advertises and does not reply. " + fact("engro_vs") +
         ". An answer, signed by the maker, in his language, is the cheapest relationship available."),
        ("Show the result, not the badge", cap(fact("no_proof")) + ". Proof on a field like his is the one argument a "
         "commodity price cannot copy."),
        ("Let him speak", cap(fact("farmer_voice")) + ". A brand that hands the farmer the microphone is remembered for it "
         "long after the day.")]):
        x = 0.6 + i * 4.1
        rect(s, x, 1.3, 3.9, 0.75, NAVY)
        txt(s, x + 0.25, 1.5, 3.4, 0.4, head, 14, True, WHITE, space=0)
        rect(s, x, 2.05, 3.9, 3.9, PALE)
        txt(s, x + 0.25, 2.3, 3.4, 3.4, body, 12, False, INK)

    # 7 emotion x proof
    s = frame(prs, "6  ·  EMOTION x PROOF", "One system, not two campaigns",
              "Salam Kissan earns the right to be believed; the proof campaign spends it. Neither works alone.",
              "Our own analysis: Case 1 (Salam Kissan) and Case 2 (the 10% promise) decks  ·  " + BRIEF)
    table(s, 0.6, 1.3, [2.6, 4.7, 4.8], [
        ["", "Emotional storytelling", "Product superiority"],
        ["What it is", "Salam Kissan: the farmer answers the country's salute, in his own voice",
         "Dus Feesad Aur: the 10% shown on a field like his, with the recipe published"],
        ["What it earns", "Standing: the right to be listened to by farmers and by the country",
         "Preference: a reason to change the bag this season"],
        ["When it runs", "December, and the replies collected through the year",
         "Sowing and harvest, when the decision is actually made"],
        ["How they hand over", "The farmer who replies in December is the farmer whose field is measured in March",
         "The result published in March becomes the story told in December"]],
        10.5, 0.95)

    # 8 measurement
    s = frame(prs, "7  ·  HOW WE WOULD MEASURE IT", "A measurement frame that can fail",
              "Each objective gets one number that can go the wrong way. Anything we cannot measure, we say so.",
              "Needs from Fatima: brand-health attribution of Kissan Day, TikTok results, the 2023 giveaway rules, "
              "demo-plot data, dealer offtake by district.")
    table(s, 0.6, 1.3, [3.0, 4.3, 4.8], [
        ["Objective", "The number", "Where it comes from"],
        ["Awareness of Salam Kissan and Kissan Day", "Prompted and spontaneous awareness of the day, and who is credited for it",
         "Brand health, with an attribution question added - it does not exist today"],
        ["Farmers heard, not just reached", "Replies collected: how many farmers, from how many districts, published each year",
         "Our own collection; the number is the movement's proof"],
        ["Belief in the 10%", "Share of farmers who say they have seen a result on a field like theirs",
         "Brand health, plus demo-plot and Ki Jeet records"],
        ["Preference and trade", "NP and CAN offtake by district against the districts we showed proof in",
         "Fatima's dealer data - the only place the campaign meets the bag"],
        ["Quality of the relationship", "Reply rate and answer time on farmer questions in comments and the app",
         "Our own logs; today the honest answer is that most questions go unanswered"]],
        10, 0.88)

    # 9 long term
    s = frame(prs, "8  ·  THE LONG TERM", "What stays the same for three years, and what changes every season",
              "Consistency is a system, not a promise: one platform, one proof engine, one calendar, one library.",
              "Our own working model  ·  " + BRIEF + " p14")
    panel(s, 0.6, 1.25, 3.9, 4.9, "What stays fixed", [
        "The two platforms: the farmer's voice, and the proof.",
        "The signature and the sign-off, so seven years of equity keep compounding instead of restarting.",
        "One annual published record: replies collected, plots measured, questions answered.",
        "The rule that a claim on air points to a named field."], 11)
    panel(s, 4.7, 1.25, 3.9, 4.9, "What changes each season", [
        "The crop, the district, the soil and the farmer in the film.",
        "The question being answered, taken from what farmers actually asked that season.",
        "The creator and partner mix, as the movement widens.",
        "The number we publish, which should get bigger or be explained."], 11)
    panel(s, 8.8, 1.25, 3.9, 4.9, "How quality stays consistent", [
        "One team on both brands, with the agronomy sign-off built into the creative route, not added after.",
        "A shared asset library: every farmer story shot once, cut for TV, for a short, for the dealer counter and for the app.",
        "A quarterly evidence review: what we claimed, what the field showed, what we stop saying.",
        "If the proof does not arrive, we say so and change the claim - before a regulator or a farmer does."], 11)


def free_path(p: Path) -> Path:
    if not p.exists():
        return p
    n = 2
    while (q := p.with_name(f"{p.stem}-v{n}{p.suffix}")).exists():
        n += 1
    return q


def main() -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    build(prs)
    missing = [k for k in USED if k not in FACTS]
    if missing:
        raise SystemExit(f"figures used with no source: {missing}")
    out = free_path(WS / "out" / "Sarsabz_point_of_view.pptx")
    prs.save(out)
    print(f"{out}  ({len(prs.slides)} slides; {len(set(USED))} sourced figures used)")


if __name__ == "__main__":
    main()
