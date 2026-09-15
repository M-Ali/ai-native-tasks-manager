"""What the Fatima Fertilizer creative-agency brief actually says, page-traced.

Source: `data/Creative Pitch Deck (1).pdf` (32 pages). Pages 16-26 are image-only
brand-guideline pages and pages 28-31 are charts from the U&A + Brand Health study;
their text was read from rendered page images, not a text layer.

Rule: every item carries the brief page it was read from. If you change an item,
re-read that page first. Do not add a requirement the brief does not state - put it in
OPEN_QUESTIONS instead.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field

BRIEF_PDF = "data/Creative Pitch Deck (1).pdf"


@dataclass(frozen=True)
class Item:
    text: str
    page: int


@dataclass(frozen=True)
class Brief:
    client: str = "Fatima Fertilizer Company Limited (JV of Fatima Group and Arif Habib)"
    title: str = "Creative Marketing Brief 2026 - creative agency pitch"
    brands: tuple[str, ...] = ("Sarsabz", "Bubber Sher (regional, North Zone)", "Pakarab CAN")
    scope_entities: tuple[str, ...] = (
        "Fatima Fertilizer Company Ltd (Sadiqabad)",
        "Pak Arab Fertilizers Ltd (Multan)",
        "FatimaFert Ltd (Sheikhupura)",
        "all agri-related businesses under Fatima Group",
    )
    presentation_minutes: int = 90
    pitch_date: str | None = None  # not stated in the brief
    budget: str | None = None  # not stated in the brief


PITCH_CONTENTS: tuple[Item, ...] = (
    Item("Introduction of the agency", 11),
    Item("Detailed agency profile incl. team structure and dedicated resources", 11),
    Item("Achievements & awards", 11),
    Item("Relevant case studies", 12),
    Item("Understanding of our brands and consumer", 12),
    Item("Solution to the case(s) considering corporate values, brand purpose, brand history", 12),
    Item("Any other suggestions for our brands (welcomed)", 12),
)

CASE1_BACKGROUND: tuple[Item, ...] = (
    Item("Salam Kissan began in 2019 to meet farmers' 'need of appreciation' in the form of Farmers Day", 12),
    Item("Celebrated on 18 December with a 360° campaign; the Government of Pakistan now officially recognises "
         "18 December as National Kissan Day", 12),
    Item("Described as 'one of Pakistan's most recognizable purpose-led initiatives'", 12),
    Item("2026 ambition: make Salam Kissan 'bigger and better', take the movement to the next level and "
         "'strengthen our ownership of the cause'", 12),
)

# The four things the Case 1 concept and 360° campaign must do (p.12).
CASE1_TASK: tuple[Item, ...] = (
    Item("Celebrate Pakistan's farmers and their contribution towards the country", 12),
    Item("Strengthen Sarsabz's position as the pioneer and champion of the Salam Kissan movement", 12),
    Item("Connect with both rural and urban audiences, especially the younger generation", 12),
    Item("Create relevance beyond a single day and transform Salam Kissan into a nationwide movement", 12),
)

# The goal: maximise awareness of Salam Kissan and National Kissan Day, rural and urban, making people understand (p.12):
CASE1_GOAL: tuple[Item, ...] = (
    Item("The importance of farmers in our daily lives", 12),
    Item("The contribution of agriculture towards Pakistan's economy", 12),
    Item("The role played by Sarsabz in celebrating and empowering the farming community", 12),
)

CASE1_CHANNELS: tuple[str, ...] = ("ATL", "BTL", "PR", "Events", "On-Ground Activations")

CASE2_TASK: tuple[Item, ...] = (
    Item("Review/audit current Sarsabz brand assets, communication and campaigns", 13),
    Item("Assess Sarsabz Ki Jeet alongside the brand's other platforms", 13),
    Item("Recommend a TACTICAL campaign on the functional aspects pushing "
         "'10 feesad se bhi ziada izafi paidawar' among farmers", 13),
)

# The nine things Case 2 explicitly asks the agency to recommend (p.13).
CASE2_RECOMMEND: tuple[Item, ...] = (
    Item("Opportunities to strengthen the brand positioning", 13),
    Item("Areas of improvement in current communication", 13),
    Item("New creative territories Sarsabz can explore", 13),
    Item("Opportunities for stronger emotional AND functional connections with farmers", 13),
    Item("Recommendations to improve brand consistency across touchpoints", 13),
    Item("Whether Sarsabz Ki Jeet farmer wins can be authentic proof for the yield promise", 13),
    Item("An identified brand or business challenge", 13),
    Item("Ways to strengthen Sarsabz's leadership position", 13),
    Item("Ways to create engagement with farmers and relevant stakeholders", 13),
)

CASE2_ROADMAP_GOALS: tuple[Item, ...] = (
    Item("Enhance brand equity and communication effectiveness", 14),
    Item("Strengthen connection between Sarsabz and the farming community", 14),
    Item("New creative territories: more relevant and differentiated", 14),
    Item("Improve brand expression, storytelling and consumer engagement", 14),
    Item("Tactical idea creating awareness, engagement and preference", 14),
    Item("Practical use of Sarsabz Ki Jeet real farmer wins as proof of yield promise", 14),
)

CASE2_CHANNELS: tuple[str, ...] = ("ATL", "BTL", "Digital", "Dealer/Trade", "On-Ground")

EXPECTATIONS: tuple[Item, ...] = (
    Item("Clear point of view on where the brands should go, and why", 14),
    Item("Strategic vision for Sarsabz AND Baber Sher", 14),
    Item("Understanding of agri audience, fertilizer category dynamics, farmer behaviour", 14),
    Item("Innovative, culturally relevant ideas across touchpoints", 14),
    Item("Balance emotional storytelling with product superiority", 14),
    Item("Practical, measurable approach to campaign effectiveness", 14),
    Item("Long-term thinking; consistent high-quality creative delivery", 14),
)

BRAND_FACTS: tuple[Item, ...] = (
    Item("Sarsabz launched 2011; tagline 'Zameen ki Jaan, Fasal ki Shaan'", 5),
    Item("Only nitrate fertilizer manufacturer; promise: avg 10% higher yield vs conventional "
         "with combined NP + CAN", 5),
    Item("Promise documented through 500+ demo plots with research institutes and govt stations", 5),
    Item("Brand pillars: Experts, Caring & farmer-focused, One-stop solution", 5),
    Item("Branded house: Sarsabz DAP, Urea, NP, CAN-G; Bubber Sher (Urea, DAP) and Pak Arab CAN "
         "are house-of-brands", 20),
    Item("Brand essence: 'Feeding the Soil, Fueling the Future. Empowering and Recognizing the Farmer.'", 25),
    Item("Brand character: Progressive & Steadfast Champion", 24),
    Item("Competitive frame: FFC Sona ('Gold', reliability/legacy); Engro Heera ('diamond', "
         "transformation; dealer discounts; consistency concerns)", 23),
    Item("Core target (guidelines): male farmers 25-55, Punjab & Sindh, 12-150+ acres, "
         ">=50% income from farming, TikTok-engaged", 23),
    Item("Target (brief body): 18-60, SEC A-lower B & C, upper D, min. secondary education", 6),
    Item("Sarsabz Pakistan App 800,000+ downloads; Sarsabz Asaan (2023) Rs.500bn sales recorded", 7),
    Item("Seasons: Rabi Oct-Mar; Kharif Apr-Sep", 6),
)

# U&A + Brand Health study, master-brand level (pp.29-31). Competitors are unnamed.
AWARENESS: dict[str, dict[str, dict[int, int]]] = {
    "Fatima": {"tom": {2013: 4, 2016: 5, 2018: 7, 2022: 7, 2026: 22},
               "spont": {2013: 52, 2016: 57, 2018: 43, 2022: 10, 2026: 81},
               "aided": {2013: 32, 2016: 38, 2018: 50, 2022: 29}},
    "Competitor 1": {"tom": {2026: 62}, "spont": {2026: 37}},
    "Competitor 2": {"tom": {2026: 16}, "spont": {2026: 61}},
}

OPEN_QUESTIONS: tuple[str, ...] = (
    "Pitch date and submission deadline (not stated)",
    "Budget / media spend envelope for the tactical campaign (not stated)",
    "Which competitors are 'Competitor 1-4' in the brand health study; sample, method, fieldwork dates",
    "Market share by product (Urea, DAP, NP, CAN) and trend - brief gives none",
    "Is the '6% ahead of competition' TOM claim vs Competitor 2 only? Competitor 1 shows 62% (p.29)",
    "2022 spontaneous awareness of 10% vs 43% in 2018 - method change or real drop? (p.31)",
    "Target age: 18-60 (p.6) or 25-55 (p.23)?",
    "Sarsabz Ki Jeet: how many Sarsabz farmers won, which years/districts, rules of the Punjab "
    "wheat yield competition, and whether results can be used in advertising",
    "Demo-plot dataset behind the 10% claim: can the agency see it (crops, seasons, spread)?",
    "Is the 10% claim substantiated for advertising-standards purposes (PEMRA / Competition Act s.10)?",
    "Bubber Sher: expected depth in the pitch - neither case covers it but p.14 expects a vision",
)


def as_dict() -> dict:
    def items(xs):
        return [asdict(x) for x in xs]

    return {
        "brief": asdict(Brief()),
        "pitch_contents": items(PITCH_CONTENTS),
        "case2": {"task": items(CASE2_TASK), "recommend": items(CASE2_RECOMMEND),
                  "roadmap": items(CASE2_ROADMAP_GOALS), "channels": list(CASE2_CHANNELS)},
        "expectations": items(EXPECTATIONS),
        "brand_facts": items(BRAND_FACTS),
        "awareness": AWARENESS,
        "open_questions": list(OPEN_QUESTIONS),
    }
