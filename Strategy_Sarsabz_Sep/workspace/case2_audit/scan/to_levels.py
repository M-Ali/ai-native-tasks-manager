"""Classify both brands' coded posts to Brand Meaning Ladder levels -> levels.csv.

Mapping rule: classify by what the claim DOES, not its topic, and put each item at the deepest
level the evidence supports and no deeper (brand-laddering/references/levels.md).

Decisions recorded here so they can be challenged:
  L1  brand_goodwill        presence: sponsorship entertainment, Ramzan, corporate PR, documentaries
  L2  product_quality       what the bag contains: bio-active zinc, blends, granule, TSP, NP+
  L2  rewards_prizes        "buy this bag and win" - a reason to buy, no meaning beyond the bag
  L2  authenticity          a verifiable feature of the product (seal, SMS check)
  L3  yield_increase        what it does for me: more maunds
  L3  advisory_service      what it does for me: my crop problem solved, my soil tested
  L3  cost_saving           what it does for me: less spent per acre
  L4  farmer_recognition    what choosing says about me: the farmer as someone worth saluting
  L5  women_empowerment     a stated belief about how things should be
  L5  national_identity     a stated belief / way of life (Pakistan, the land)
"""
import csv
from pathlib import Path

HERE = Path(__file__).parent
LEVEL = {
    "brand_goodwill": (1, "presence: sponsorship, festive or corporate film; no product or belief claim"),
    "product_quality": (2, "attribute: what is in the bag"),
    "rewards_prizes": (2, "attribute-level reason to buy: scratch code and prize draw"),
    "authenticity": (2, "attribute: a verifiable seal on the bag"),
    "yield_increase": (3, "benefit: more yield per acre"),
    "advisory_service": (3, "benefit: my crop problem solved, my soil tested"),
    "cost_saving": (3, "benefit: less cost per acre, more profit"),
    "farmer_recognition": (4, "psychological: what being a farmer, and choosing this brand, says about me"),
    "women_empowerment": (5, "stated belief about women's role in farming"),
    "national_identity": (5, "stated belief / way of life: the land and the nation"),
    "other": (1, "no identifiable claim"),
}
SRC = {"Sarsabz": HERE.parent / "owned" / "codes_window.csv",
       "Engro": HERE / "engro" / "codes_window.csv"}

out = []
for brand, path in SRC.items():
    for r in csv.DictReader(path.open(encoding="utf-8")):
        lvl, why = LEVEL[r["claim_primary"]]
        out.append({"post_id": r["video_id"], "brand": brand, "ladder_level": lvl,
                    "ladder_reason": f"{r['claim_primary']}: {why}",
                    "platform": r["platform"], "views": r["views"], "title": r["title"]})

with (HERE / "levels.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(out[0]))
    w.writeheader()
    w.writerows(out)
print(f"{len(out)} rows -> levels.csv")
