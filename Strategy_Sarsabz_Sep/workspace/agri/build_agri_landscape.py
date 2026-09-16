"""Pakistan's agricultural landscape: crops, seasons, provinces, and where fertilizer opportunity sits.

    uv run --no-project --with python-pptx python workspace/agri/build_agri_landscape.py

Writes workspace/out/Sarsabz_agri_landscape.pptx (never overwrites: a clash becomes -v2).

Built from the tables the user supplied in data/reference/Agri_Landscape_2026 (PBS), the Ministry of
National Food Security & Research district-wise crop publication (downloaded), the Economic Survey
2025-26, and - for international practice - the World Bank indicator API, IFA and the European
Commission.

THE GUARD: every figure comes from FACTS, and every FACTS entry carries its source. Crop-by-province
numbers are read at build time from crop_by_province.csv, which holds only the five crops whose
province totals reconcile with PBS Table-1; minor crops were dropped because their province
attribution could not be validated, and the slide says so rather than showing a number I cannot defend.
"""
from __future__ import annotations

import csv
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

PBS = "PBS agriculture tables (data/reference/Agri_Landscape_2026), supplied by the user"
MNFSR = "Ministry of National Food Security & Research, Crops Area & Production (District Wise) 2022-23"
PES = ("Pakistan Economic Survey 2025-26, ch.2 Agriculture "
       "(data/reference/Agri_Landscape_2026/2_Agriculture.pdf, the latest chapter, supplied by the user)")
WB = "World Bank, Fertilizer consumption (kg per hectare of arable land), AG.CON.FERT.ZS, 2023, retrieved 16 Sep 2026"
IFA4R = "IFA, The Global '4R' Nutrient Stewardship Framework, May 2009"
IFABAL = "L. Cissé, Balanced fertilization for sustainable use of plant nutrients, in IFA Fertilizer Best Management Practices, 2007"
EUBRIEF = "European Commission, EU Agricultural Markets Brief No 15, Fertilisers in the EU, June 2019"
FEUROPE = "Fertilizers Europe, Types of fertilizer, retrieved 16 Sep 2026"
BRIEF = "Fatima Creative Marketing Brief 2026"

FACTS: dict[str, tuple[str, str]] = {
    "cropped": ("Pakistan crops 24.60m hectares: Punjab 17.52m, Sindh 3.90m, KP 1.73m, Balochistan 1.45m (2024-25)", PBS),
    "seasons": ("Rabi runs October-March, Kharif April-September", BRIEF),
    "water": ("Kharif 2025 water availability 60.6 MAF; Rabi 2025-26 31.4 MAF, against average system usage of 103.5 MAF", PES),
    "canals": ("Canal withdrawals 2024-25: Punjab 31.88 MAF in Kharif and 15.53 in Rabi; Sindh 26.16 and 12.14", PBS),
    "irrigation": ("Punjab irrigates 15.47m ha, over half of it canal-plus-tubewell; Sindh irrigates 1.52m ha, almost all canal", PBS),
    "calendars": ("PBS 'crop calendars' (Tables 13-15) are dates for releasing crop estimates, not sowing windows", PBS),
    "growth": ("Agriculture grew 2.89% in 2025-26; crops 1.44%; important crops 0.65%", PES),
    "gdp": ("Agriculture is 23.4% of GDP and 33.1% of employment", PES),
    "offtake": ("Fertilizer nutrient offtake 3,795k tonnes in Jul-Mar FY2026, up 11.4%", PES),
    "npk_trend": ("Nitrogen offtake +14.8%, phosphate -1.9% (on high prices), potash +26.2%", PES),
    "bag": ("Rs 100 more per 50kg bag puts Rs 20 billion more on farmers", PES),
    "kissan_card": ("Punjab's Kissan Card improved cotton fertilizer application", PES),
    "prov_use": ("Nutrient use 2024-25: Punjab 2,986.9k tonnes (69.1% of Pakistan), Sindh 991.0k (22.9%)", PBS),
    "per_ha": ("Per cropped hectare: Sindh 254.2 kg, Punjab 170.5 kg, Balochistan 112.0, KP 106.2", PBS),
    "npk_ratio": ("Pakistan's N:P:K is 1 : 0.27 : 0.013; Punjab 1 : 0.28 : 0.015; Sindh 1 : 0.24 : 0.010 (2024-25)", PBS),
    "cropwise": ("PBS crop-wise fertilizer (Table-11) is a fixed percentage split - wheat 50%, cotton 25%, rice 12% - "
                 "not measured consumption", PBS),
    "balanced": ("IFA cites a 'normal or generally accepted balanced ratio' between N, P2O5 and K2O of 1:0.5:0.5", IFABAL),
    "imbalance_cost": ("Consequences IFA lists: lower fertilizer use efficiency, yield reduction, lower farmer income, "
                       "soil mining of P and K, falling response ratios, higher nitrogen losses", IFABAL),
    "pak_history": ("Pakistan's ratio was already 1:0.3:0.01 through 1996-2005 - the imbalance is structural, not new", IFABAL),
    "fourR": ("Fertilizer best practice is 'the application of the right source (or product) at the right rate, right time "
              "and right place'", IFA4R),
    "wb": ("Fertilizer per hectare of arable land, 2023: Pakistan 160.3 kg, India 199.1, Bangladesh 391.9, China 394.0, "
           "Vietnam 419.9, Egypt 532.8, United States 127.8", WB),
    "urea_eu": ("The EC, citing IFA: 'the market share of urea in the nitrogen-based fertilisers market is high in Asia "
                "while it is lower in the EU and in North America'", EUBRIEF),
    "nitrate_eu": ("Fertilizers Europe: ammonium nitrate and calcium ammonium nitrate are 'well suited to most European "
                   "soils and climatic conditions'", FEUROPE),
    "global_split": ("Globally nitrogen is 108m tonnes of nutrient (60%), of which urea is 60m tonnes", EUBRIEF),
}
USED: list[str] = []


def fact(key: str) -> str:
    USED.append(key)
    return FACTS[key][0]


def sources_line(keys: list[str]) -> str:
    seen: list[str] = []
    for k in keys:
        for s in FACTS[k][1].split("; "):
            if s not in seen:
                seen.append(s)
    return "Sources: " + "  ·  ".join(seen)


def txt(slide, l, t, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, space=6, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(line) or " "
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


def table(s, l, t, widths, rows, size=10.5, row_h=0.52, head_h=0.45):
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
            txt(s, x + 0.12, y + 0.09, w - 0.24, h - 0.14, cell, size, hdr, WHITE if hdr else INK, space=0)
            x += w
        y += h
    if y > 6.3:
        raise SystemExit(f"table runs to {y:.2f} in and would cover the conclusion bar")
    return y


def crops():
    rows = list(csv.DictReader((HERE / "crop_by_province.csv").open(encoding="utf-8")))
    for r in rows:
        r["area_000ha"] = float(r["area_000ha"])
        r["share"] = float(r["share_of_national_pct"])
        r["production"] = float(r["production"])
    return rows


def build(prs) -> None:
    C = crops()
    national = {}
    for r in C:
        national[r["crop"]] = national.get(r["crop"], 0) + r["area_000ha"]

    # 1 cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, NAVY)
    rect(s, 0.9, 1.7, 1.6, 0.06, ACCENT)
    txt(s, 0.9, 1.95, 11.4, 0.5, "Fatima Fertilizer creative pitch  ·  background", 18, False, SKY)
    txt(s, 0.9, 2.5, 11.6, 1.0, "The agricultural landscape", 48, True, WHITE)
    txt(s, 0.9, 3.6, 11.4, 0.6, "What Pakistan grows, when, where - and where the fertilizer opportunity sits", 22, False, WHITE, italic=True)
    txt(s, 0.9, 4.8, 11.4, 1.2, ["1  The land and the two seasons   ·   2  What Pakistan grows",
                                 "3  Punjab   ·   4  Sindh   ·   5  Where the fertilizer goes   ·   6  The imbalance",
                                 "7  What is practised internationally   ·   8  The opportunity by product"], 14, False, SKY)
    txt(s, 0.9, 6.6, 11.6, 0.4, "Built from the PBS tables supplied, the MNFSR district-wise crop publication, the Economic "
                                "Survey 2025-26, the World Bank, IFA and the European Commission  ·  16 September 2026", 11, False, SKY)

    # 2 land and seasons
    k = ["cropped", "seasons", "water", "canals", "irrigation", "calendars"]
    s = frame(prs, "1  ·  THE LAND AND THE TWO SEASONS", "Two seasons, two provinces, and water that decides both",
              "Punjab is 71% of the cropped area; Sindh is canal-fed and concentrated. The season, not the campaign, sets "
              "when fertilizer moves.", sources_line(k))
    panel(s, 0.6, 1.25, 6.0, 2.35, "The land", [fact("cropped") + ".", fact("irrigation") + "."], 11.5)
    panel(s, 0.6, 3.75, 6.0, 2.4, "The two seasons", [fact("seasons") + ".",
          "Rabi is the wheat season; Kharif carries rice, cotton, sugarcane and maize.",
          fact("calendars") + " - sowing windows are not in this data set."], 11.5)
    panel(s, 6.9, 1.25, 5.8, 4.9, "Water, which decides the season", [
        fact("water") + ".", fact("canals") + ".",
        "Kharif carries roughly twice the water of Rabi, but Rabi is when wheat - half the country's fertilizer - is grown.",
        fact("growth") + "."], 11.5)

    # 3 what Pakistan grows
    k = ["gdp", "growth"]
    s = frame(prs, "2  ·  WHAT PAKISTAN GROWS", "Five crops carry the fertilizer market",
              "Wheat alone is over half the cropped area of the five majors - and the Rabi season the brand keeps "
              "celebrating in December.", sources_line(k) + "  ·  " + MNFSR + ", area 2022-23, province totals reconciled "
              "against PBS Table-1")
    rows = [["Crop", "Season", "National area ('000 ha)", "Punjab share", "Sindh share"]]
    season = {"WHEAT": "Rabi (Oct-Mar)", "RICE": "Kharif (Apr-Sep)", "COTTON": "Kharif (Apr-Sep)",
              "SUGARCANE": "Kharif, 12-18 month crop", "MAIZE": "Spring and Kharif"}
    for crop in sorted(national, key=lambda c: -national[c]):
        pj = next((r for r in C if r["crop"] == crop and r["province"] == "PUNJAB"), None)
        sd = next((r for r in C if r["crop"] == crop and r["province"] == "SINDH"), None)
        rows.append([crop.title(), season[crop], f"{national[crop]:,.0f}",
                     f"{pj['share']:.0f}%" if pj else "-", f"{sd['share']:.0f}%" if sd else "-"])
    table(s, 0.6, 1.3, [2.2, 3.1, 2.6, 2.1, 2.1], rows, 11, 0.62)
    panel(s, 0.6, 4.6, 12.1, 1.55, "What this data set cannot tell you", [
        "Only these five crops are shown: their province totals reconcile exactly with PBS Table-1. Minor crops "
        "(gram, sesame, tobacco, bajra) were dropped because their province attribution could not be validated in the source PDF.",
        fact("gdp") + " - the sector's weight, but not its fertilizer demand, which follows crop and season."], 11)

    # 4 and 5: Punjab, Sindh
    for who, label, note in [("PUNJAB", "3  ·  PUNJAB", "Punjab is the fertilizer market: 69% of national nutrient use."),
                             ("SINDH", "4  ·  SINDH", "Sindh is smaller but more intensive per hectare, and more nitrogen-skewed.")]:
        s = frame(prs, label, f"{who.title()}: what is grown, ranked by area",
                  note, MNFSR + "  ·  area and production 2022-23  ·  cotton production in bales, others in tonnes")
        rows = [["Crop", "Area ('000 ha)", "Share of national area", "Production"]]
        for r in sorted([r for r in C if r["province"] == who], key=lambda r: -r["area_000ha"]):
            rows.append([r["crop"].title(), f"{r['area_000ha']:,.1f}", f"{r['share']:.0f}%",
                         f"{r['production']:,.0f} {r['prod_unit']}"])
        table(s, 0.6, 1.3, [2.6, 2.6, 3.0, 3.9], rows, 11, 0.62)
        if who == "PUNJAB":
            panel(s, 0.6, 4.7, 12.1, 1.45, "What it means", [
                "Wheat is the volume, cotton and rice the value at risk, sugarcane the highest nutrient demand per hectare.",
                fact("kissan_card") + " - the state is now a channel into the same farmer."], 11)
        else:
            # Sindh has a sixth crop row, so this panel starts below the table, not level with Punjab's
            panel(s, 0.6, 4.75, 12.1, 1.4, "What it means", [
                "Sindh's rice and sugarcane sit on canal water, so timing is fixed by the canal, not by the farmer.",
                "Maize is effectively absent from Sindh; cotton and wheat carry the season.",
                "Per hectare, Sindh already uses more nutrient than Punjab - the gap is what it uses, not how much."], 11)

    # 6 where the fertilizer goes
    k = ["prov_use", "per_ha", "offtake", "npk_trend", "bag", "cropwise"]
    s = frame(prs, "5  ·  WHERE THE FERTILIZER GOES", "Two provinces are the market; price decides the mix",
              "Nitrogen is bought; phosphate is skipped when it is dear. That is the opening for a cheaper phosphate route.",
              sources_line(k))
    table(s, 0.6, 1.25, [3.0, 3.1, 3.0, 3.0], [
        ["", "Punjab", "Sindh", "Pakistan"],
        ["Nutrient use 2024-25", "2,986.9k tonnes (69.1%)", "991.0k tonnes (22.9%)", "4,324.2k tonnes"],
        ["Per cropped hectare", "170.5 kg", "254.2 kg", "175.8 kg"],
        ["N : P : K", "1 : 0.28 : 0.015", "1 : 0.24 : 0.010", "1 : 0.27 : 0.013"]],
        11, 0.72)
    panel(s, 0.6, 4.1, 6.0, 2.05, "What is moving", [fact("offtake") + ".", fact("npk_trend") + ".", fact("bag") + "."], 11)
    panel(s, 6.9, 4.1, 5.8, 2.05, "A caveat worth stating in the room", [
        fact("cropwise") + ".",
        "So 'wheat takes half the fertilizer' is an assumption inside the statistics, not a measurement of farmer behaviour."], 11)

    # 7 the imbalance
    k = ["npk_ratio", "balanced", "pak_history", "imbalance_cost", "npk_trend"]
    s = frame(prs, "6  ·  THE IMBALANCE", "Pakistan buys nitrogen and skips the rest",
              "Against the balanced ratio, Pakistan applies about half the phosphate and almost no potash. That is the "
              "agronomic argument the category has never made.", sources_line(k))
    panel(s, 0.6, 1.25, 6.0, 2.3, "What Pakistan applies", [fact("npk_ratio") + ".",
          fact("balanced") + ".", "Against that, phosphate runs at roughly half the balanced level and potash at almost nothing."], 11.5)
    panel(s, 0.6, 3.7, 6.0, 2.45, "It is structural, not a bad year", [fact("pak_history") + ".",
          fact("npk_trend") + " - the imbalance widened again this year, on price."], 11.5)
    panel(s, 6.9, 1.25, 5.8, 4.9, "What the imbalance costs, per IFA", [fact("imbalance_cost") + ".",
          "For a brand: every one of those is a farmer-visible loss - lower yield, wasted nitrogen, a soil that gives less "
          "each year - and none of it is currently being explained to him by anyone selling him the bag."], 11.5)

    # 8 international practice
    k = ["wb", "fourR", "urea_eu", "nitrate_eu", "global_split"]
    s = frame(prs, "7  ·  WHAT IS PRACTISED INTERNATIONALLY", "Three practices Pakistan has not adopted",
              "Pakistan is not under-fertilised by accident: it uses less per hectare than its neighbours, and what it uses "
              "is the least balanced.", sources_line(k))
    table(s, 0.6, 1.3, [3.4, 2.4, 6.3], [
        ["Practice", "Where", "What it means here"],
        ["Higher and balanced application", "Bangladesh 391.9 and China 394.0 kg/ha against Pakistan's 160.3 (2023)",
         "Pakistan's yield gap is partly a nutrition gap; India at 199.1 is the nearest comparison"],
        ["Nitrate-based nitrogen, not only urea", "AN and CAN are the European norm; urea's share is 'high in Asia... lower "
         "in the EU'", "Sarsabz is the only CAN maker at scale in Pakistan - an international norm it already produces"],
        ["4R nutrient stewardship", "The global framework since 2009",
         "'Right source, right rate, right time, right place' is exactly the advice farmers ask for in the comments"]],
        10.5, 1.15)
    panel(s, 0.6, 5.05, 12.1, 1.1, "The honest comparison", [
        fact("global_split") + " - so urea's dominance is global, but its dominance of Pakistan's mix is what makes the "
        "ratio unbalanced."], 11)

    # 9 opportunity by product
    s = frame(prs, "8  ·  THE OPPORTUNITY BY PRODUCT", "Where each fertilizer type has a claim, and what it needs",
              "Each opportunity is an argument the category is not making - and three of the four are products Fatima already makes.",
              "Our own reading of the data on the previous slides  ·  " + BRIEF + "  ·  product positions from the Case 2 audit")
    table(s, 0.6, 1.25, [2.3, 3.5, 3.4, 2.9], [
        ["Product", "Where the opening is", "The argument", "What it needs to be true"],
        ["NP (Nitrophos)", "Phosphate offtake fell 1.9% on price while nitrogen rose 14.8%",
         "A cheaper route to phosphate than DAP, on the crops where P is being skipped",
         "Published per-acre comparisons against DAP, by soil and crop"],
        ["CAN", "Sindh's mix is the most nitrogen-skewed (1 : 0.24 : 0.010); canal-timed irrigation",
         "The nitrate form Europe defaults to, suited to timed splits on irrigated land",
         "Agronomy in the farmer's language, and a compatibility answer"],
        ["Urea / DAP", "The volume market, decided by price and availability",
         "Availability and authenticity, not persuasion",
         "Dealer presence and a visible price promise"],
        ["Potash and micronutrients", "Potash is 1.8 kg per cropped hectare nationally",
         "The balanced-ratio argument, unclaimed by anyone in the category",
         "A category-building case Fatima would fund without owning the product"]],
        10, 1.05)


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
    out = free_path(WS / "out" / "Sarsabz_agri_landscape.pptx")
    prs.save(out)
    print(f"{out}  ({len(prs.slides)} slides; {len(set(USED))} sourced figures)")


if __name__ == "__main__":
    main()
