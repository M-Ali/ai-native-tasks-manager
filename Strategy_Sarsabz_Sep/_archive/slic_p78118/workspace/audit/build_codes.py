"""Author codes.csv for the five new dimensions, from the contact sheets + captions.

    uv run python workspace/audit/build_codes.py

Rules are matched on a distinctive caption substring rather than a transcribed
shortcode: shortcodes contain I/l/1 lookalikes and hand-copying them is the obvious
place to introduce a silent mismatch. Matching on copy the ad actually carries is both
safer and auditable - you can read a rule and check it against the post.

Anything no rule matches is written `unclear` and reported, so the gap is visible
rather than papered over.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

AUDIT = Path(__file__).parent
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FIELDS = ["product_focus", "audience_generation", "content_type", "who_is_in_frame",
          "production_value", "ecosystem_push", "campaign_platform", "language_lead"]


def C(focus, gen, ctype, frame, prod, eco="no", plat="", lang="english_first"):
    return dict(zip(FIELDS, [focus, gen, ctype, frame, prod, eco, plat, lang]))


# brand -> [(caption substring, codes)]. First match wins; order matters.
RULES: dict[str, list[tuple[str, dict]]] = {

"State Life Insurance Corporation of Pakistan": [
 ("Eid Milad-un-Nabi",      C("brand","mixed","calendar_topical","none_people","template_graphic",plat="",lang="equal")),
 ("Qimam Fellowship",       C("brand","genz","corporate_pr","staff","event_photo")),
 ("Elevate 2026",           C("brand","genx","corporate_pr","executive","event_photo")),
 ("79th Independence Day",  C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("May every child dream",  C("brand","mixed","calendar_topical","customer","studio",plat="Har Khwab Rahe Azaad")),
 ("There's a lot more for the insurance industry", C("brand","genx","corporate_pr","executive","event_photo")),
 ("E-Kachehri",             C("E-Kachehri","mixed","product_feature","none_people","template_graphic",lang="equal")),
 ("You can't always predict when you'll need care", C("brand","geny","brand_building","none_people","studio",plat="Plan Ahead With State Life")),
 ("The best opportunities are easier to seize",     C("brand","geny","brand_building","customer","studio",plat="Plan Ahead With State Life")),
 ("Back-to-school season",  C("brand","geny","brand_building","customer","studio",plat="Plan Ahead With State Life")),
 ("called on Rana Tan",     C("brand","genx","corporate_pr","executive","event_photo")),
 ("five decades of purpose",C("brand","mixed","brand_building","none_people","template_graphic")),
],

"EFU Life Assurance": [
 ("foremost commitment is to safeguard", C("brand","genx","corporate_pr","executive","studio")),
 ("Freshly made patties",   C("PRIMUS","genz","tactical_promo","none_people","stock",eco="yes")),
 ("Financial learning just got more rewarding", C("Thrive","genz","product_feature","none_people","template_graphic",eco="yes")),
 ("Some wins don't come with a big celebration", C("EFU Life WIN","genz","brand_building","none_people","template_graphic",eco="yes")),
 ("Let your confidence glow", C("PRIMUS","geny","tactical_promo","none_people","stock",eco="yes")),
 ("Prioritize your health with timely check-ups", C("PRIMUS","geny","tactical_promo","none_people","stock",eco="yes")),
 ("Leadership development aur people development", C("Nexus Pro","genx","corporate_pr","none_people","template_graphic",lang="urdu_first")),
 ("mHealth is expanding its network", C("EFU Life mHealth","genx","product_feature","none_people","template_graphic",eco="yes")),
 ("KraveMart",              C("EFU Life WIN","genz","tactical_promo","none_people","template_graphic",eco="yes")),
 ("Champions of Change",    C("brand","mixed","csr","staff","event_photo")),
 ("fear of being judged",   C("brand","geny","brand_building","expert","studio")),
 ("Download Thrive",        C("Thrive","genz","product_feature","none_people","template_graphic",eco="yes")),
],

"Jubilee Life Insurance": [
 ("Every claim tells a story of trust", C("brand","mixed","brand_building","customer","studio",plat="Befiker Mustaqbil")),
 ("Fitness and health should be about more", C("Jubilee Active","genz","product_feature","customer","ugc",eco="yes")),
 ("Iqra University",        C("brand","genz","recruitment","none_people","template_graphic")),
 ("Your steps can do more", C("Jubilee Active","genz","product_feature","customer","ugc",eco="yes")),
 ("Grill. Chill. Repeat",   C("Jubilee Active","genz","tactical_promo","none_people","stock",eco="yes")),
 ("journey begins now on WhatsApp", C("WhatsApp service","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Shifa International",    C("Jubilee Active","geny","tactical_promo","none_people","template_graphic",eco="yes")),
 ("Walk more, earn points", C("Jubilee Active","genz","product_feature","customer","ugc",eco="yes")),
 ("Iqra Labs",              C("Jubilee Active","geny","tactical_promo","none_people","template_graphic",eco="yes")),
 ("Like Reply",             C("Jubilee Active","genz","ugc_repost","customer","ugc",eco="yes")),
],

"Adamjee Life Assurance": [
 ("Managing your claims and policy", C("Adamjee Life App","geny","product_feature","none_people","studio",eco="yes")),
 ("new Customer Care App enables",   C("Adamjee Life App","geny","product_feature","customer","studio",eco="yes")),
 ("Missed appointments",             C("Adamjee Life App","geny","product_feature","customer","studio",eco="yes")),
 ("Skip the queues and pay with ease",C("Adamjee Life App","genx","product_feature","customer","studio",eco="yes",plat="Kabhi bhi Kahin bhi")),
 ("A Journey Completed",             C("brand","genz","recruitment","none_people","template_graphic")),
 ("Ten years is more than a milestone", C("brand","mixed","corporate_pr","executive","event_photo")),
 ("Everything you need, all in one place", C("Adamjee Life App","geny","product_feature","none_people","template_graphic",eco="yes",plat="Kabhi bhi Kahin bhi")),
 ("Some moments are more than celebrations", C("brand","mixed","corporate_pr","staff","event_photo")),
 ("From celebrating our freedom",    C("brand","mixed","calendar_topical","staff","event_photo")),
 ("A day of celebration",            C("brand","mixed","corporate_pr","staff","event_photo")),
 ("Hunar Founda",                    C("brand","mixed","csr","staff","event_photo")),
 ("every shade of this land",        C("brand","mixed","calendar_topical","none_people","studio",plat="The Home to Every Life")),
],

"IGI Life Insurance": [
 ("most productive thing you can do is pause", C("IGI Life Vitality","geny","product_feature","customer","stock")),
 ("Hepatitis Awareness",    C("brand","mixed","csr","staff","event_photo")),
 ("Life doesn't stay the same", C("brand","geny","brand_building","customer","studio")),
 ("who's most likely to",   C("brand","genz","corporate_pr","staff","event_photo")),
 ("20-20-20 rule",          C("IGI Life Vitality","geny","product_feature","customer","stock")),
 ("child's education doesn't have to wait", C("brand","geny","brand_building","none_people","template_graphic",plat="Life Insurance Myth")),
 ("Different faces",        C("brand","mixed","calendar_topical","staff","event_photo")),
 ("Celebrating the spirit, strength and pride", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Hot weather can make food spoil", C("IGI Life Vitality","geny","product_feature","none_people","stock")),
 ("Think life insurance is only for when you have a family", C("brand","genz","brand_building","none_people","template_graphic",plat="Life Insurance Myth")),
 ("Small, healthy choices",  C("IGI Life Vitality","genz","tactical_promo","none_people","stock",eco="yes")),
 ("too young for life insurance", C("brand","genz","brand_building","none_people","template_graphic",plat="Life Insurance Myth")),
],

"Askari Life Assurance": [
 ("One decision today can make every moment meaningful", C("brand","geny","brand_building","customer","studio",plat="Iraada Karo / Jeeto Har Ghari")),
 ("Every goal deserves a plan that grows with you", C("Golden Path","geny","product_feature","none_people","template_graphic")),
 ("lasting financial confidence", C("Golden Path","geny","product_feature","none_people","template_graphic")),
 ("Har khwab ki apni ek kahani", C("brand","geny","brand_building","none_people","studio",plat="Iraada Karo / Jeeto Har Ghari",lang="urdu_first")),
 ("Every big dream begins with a small intention", C("brand","geny","brand_building","customer","studio",plat="Iraada Karo / Jeeto Har Ghari")),
 ("Askari Life Family Takaful", C("Family Takaful","geny","product_feature","none_people","studio")),
 ("Smart choices today",    C("Golden Path","geny","product_feature","customer","studio")),
 ("can't predict every change in life", C("brand","genz","brand_building","customer","studio",plat="Iraada Karo / Jeeto Har Ghari")),
 ("Every child deserves a future", C("Education Plan","geny","product_feature","customer","studio")),
 ("child's dream should never pause", C("Education Plan","geny","product_feature","customer","studio")),
],

"TPL Life Insurance": [
 ("From Dubai to London",   C("Globewell","affluent_na","product_feature","none_people","studio")),
 ("vibrant hues of our cultural heritage", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("strategic partnership between TPL Life", C("brand","genx","corporate_pr","executive","event_photo")),
 ("round-the-clock assis",  C("Globewell","geny","product_feature","customer","studio")),
 ("successful first half",  C("brand","genx","corporate_pr","executive","event_photo")),
 ("Quality healthcare should be accessible no matter where", C("Globewell","geny","product_feature","customer","studio")),
 ("Planning your next getaway", C("Globewell","geny","product_feature","customer","studio")),
 ("travelling for work or leisure", C("Globewell","geny","product_feature","customer","studio")),
 ("Sehat Zindagi's affordable plan", C("Sehat Zindagi","mixed","product_feature","customer","studio")),
 ("Quality healthcare doesn't have to be expensive", C("Sehat Zindagi","mixed","product_feature","expert","studio")),
 ("May this Eid bring peace", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("A world of care in one plan", C("Globewell","geny","product_feature","customer","studio")),
],

"Pak-Qatar Family Takaful": [
 ("Your employees are more than just a workforce", C("Group Health Takaful","genx","product_feature","none_people","stock")),
 ("Retirement deserves",    C("Lifetime Kafalat Plan","genx","product_feature","customer","studio")),
 ("Small steps today",      C("Mahana Bachat","geny","product_feature","customer","studio")),
 ("Build Tomorrow. From Today", C("Voluntary Pension Scheme","genx","product_feature","customer","studio")),
 ("ICMA Pakistan Future of Finance", C("brand","genx","corporate_pr","executive","event_photo")),
 ("Another milestone",      C("Pension Fund Manager","genx","corporate_pr","none_people","template_graphic")),
 ("تنخواہ",                  C("Mahana Bachat","geny","product_feature","none_people","template_graphic",lang="urdu_first")),
 ("Hepatitis is silent",    C("brand","mixed","csr","none_people","template_graphic")),
 ("Bancatakaful partner",   C("brand","genx","corporate_pr","executive","event_photo")),
 ("What if your savings paid you every month", C("Mahana Bachat","geny","product_feature","customer","studio")),
],

"Dawood Family Takaful": [
 ("Sarmaya Takaful",        C("Sarmaya Takaful Plan","geny","product_feature","customer","stock")),
 ("Takaful offers more than just life protection", C("brand","geny","brand_building","none_people","template_graphic")),
 ("Arshad Nadeem",          C("brand","mixed","calendar_topical","celebrity","template_graphic")),
 ("Faith and planning go hand in hand", C("brand","mixed","brand_building","none_people","template_graphic")),
 ("Youm-e-Takbeer",         C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Samar Takaful",          C("Samar Takaful","geny","product_feature","none_people","template_graphic")),
 ("Takaful is different from conventional insurance", C("brand","mixed","brand_building","none_people","template_graphic")),
 ("Managing your Takaful contributions", C("brand","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("every small step counts", C("Sahulat Plan","geny","product_feature","none_people","template_graphic")),
 ("salute the brave souls", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("جمعہ مبارک",              C("brand","mixed","calendar_topical","none_people","template_graphic",lang="urdu_first")),
 ("tensions rise along the borders", C("brand","mixed","calendar_topical","none_people","template_graphic")),
],

"EFU Life Window Takaful": [
 ("onboarding PMDC-certified", C("EFU Life mHealth","genx","product_feature","none_people","template_graphic",eco="yes")),
 ("A nation born from a dream", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("comprehensive range of healthcare", C("EFU Life mHealth","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("among the first entities partnering", C("Federal Hemayah Pension Fund","genx","corporate_pr","none_people","template_graphic")),
 ("HEAVY RAIN ADVISORY",    C("brand","mixed","csr","none_people","template_graphic")),
 ("EFU LifeBot",            C("EFU LifeBot","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Pakistan Insurance Day", C("brand","genx","corporate_pr","executive","studio")),
 ("Hemayah Retirement Solutions", C("Hemayah Retirement","genx","product_feature","customer","studio")),
 ("Retirement planning... abhi se", C("Hemayah Retirement","geny","brand_building","customer","studio",lang="urdu_first")),
 ("only think about tax",   C("Hemayah Retirement","geny","product_feature","none_people","template_graphic")),
 ("Episode 2",              C("Hemayah Takaful Podcast","genx","product_feature","expert","studio")),
],

"Al Meezan Investments": [
 ("FY27 has begun",         C("Market research","affluent_na","brand_building","executive","studio")),
 ("Rapid Reads",            C("Market research","affluent_na","brand_building","none_people","template_graphic")),
 ("halal economy begins with conversations", C("brand","genx","corporate_pr","executive","event_photo")),
 ("Money Matters Expo in Gilgit", C("brand","mixed","corporate_pr","staff","event_photo")),
 ("first step towards a rewarding career", C("brand","genz","recruitment","none_people","template_graphic")),
 ("Durood Sharif",          C("brand","mixed","calendar_topical","none_people","template_graphic",lang="equal")),
 ("Easypaisa",              C("Meezan Funds Online","genz","product_feature","none_people","template_graphic",eco="yes")),
 ("dream home may be a future goal", C("Home Builder Plan","geny","product_feature","none_people","template_graphic")),
 ("Think retirement planning can wait", C("Meezan Tahaffuz Pension","genz","brand_building","none_people","template_graphic",plat="Myth & Fact")),
 ("Silver P",               C("brand","mixed","corporate_pr","none_people","template_graphic")),
 ("RAAST Investment ID is just a few taps", C("Meezan Funds Online","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Generate your RAAST Investment ID", C("Meezan Funds Online","geny","product_feature","none_people","template_graphic",eco="yes")),
],

"UBL Fund Managers": [
 ("Al-Ameen Funds App",     C("Al-Ameen Funds App","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Invest Via RAAST",       C("UBL Funds App","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Simplified process to start investing", C("UBL Funds App","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("blessed occasion",       C("brand","mixed","calendar_topical","none_people","template_graphic",lang="equal")),
 ("Building stronger pathways", C("brand","genx","corporate_pr","executive","event_photo")),
 ("Connecting, learning and growing together", C("brand","mixed","corporate_pr","staff","event_photo")),
 ("spirit of freedom and unity", C("brand","mixed","calendar_topical","staff","event_photo")),
 ("Bronze Sponsor",         C("brand","mixed","corporate_pr","none_people","template_graphic")),
 ("We are Hiring",          C("brand","genz","recruitment","none_people","template_graphic")),
],
}

# The grid moves between collection runs: a brand posting in between pushes an older post
# off the 12-post window, so the enriched pull and the original capture can differ by a
# row or two. These were in the original collection and are coded from the contact sheet.
EXTRA = {
 "DcFqcCxEgUF": C("brand", "mixed", "csr", "none_people", "template_graphic"),
}

# Posts with no usable caption still get coded from the contact sheet, keyed by date+brand.
BY_IMAGE = {
 ("Pak-Qatar Family Takaful", "2026-08-25"): C("brand","mixed","corporate_pr","staff","event_photo"),
 ("Pak-Qatar Family Takaful", "2026-07-24"): C("brand","geny","product_feature","customer","studio"),
 ("EFU Life Window Takaful",  "2026-07-10"): C("Hemayah x BankIslami","mixed","corporate_pr","none_people","template_graphic"),
 ("UBL Fund Managers",        ""):           C("UBL Funds App","affluent_na","product_feature","none_people","template_graphic",eco="yes"),
}


def norm(s: str) -> str:
    """Fold the typography brands actually use into something matchable.

    Two real cases from this sample: curly apostrophes everywhere, and one advertiser
    writing whole captions in mathematical-bold Unicode (U+1D400 block), which looks
    like Latin text and matches none of it. NFKC folds those to ASCII.
    """
    import unicodedata
    s = unicodedata.normalize("NFKC", s)
    for a, b in (("’", "'"), ("‘", "'"), ("“", '"'), ("”", '"'),
                 ("–", "-"), ("—", "-")):
        s = s.replace(a, b)
    return s.lower()


def main() -> None:
    rows = list(csv.DictReader((AUDIT / "enriched" / "posts.csv").open(encoding="utf-8")))
    out, unmatched = [], []
    for r in rows:
        brand, cap, date = r["brand"], r.get("caption") or "", r.get("date") or ""
        cap = norm(cap)
        codes = None
        for needle, c in RULES.get(brand, []):
            if norm(needle) in cap:
                codes = c
                break
        if codes is None:
            codes = BY_IMAGE.get((brand, date))
        if codes is None:
            unmatched.append((brand, date, cap[:60]))
            codes = C("", "unclear", "", "none_people", "unclear")
        # audience_generation must stay inside the vocabulary; affluent is not an age band
        g = codes["audience_generation"]
        row = {"post_id": r["shortcode"], **codes}
        row["audience_generation"] = "genx" if g == "affluent_na" else g
        out.append(row)

    for pid, codes in EXTRA.items():
        if not any(o["post_id"] == pid for o in out):
            g = codes["audience_generation"]
            row = {"post_id": pid, **codes}
            row["audience_generation"] = "genx" if g == "affluent_na" else g
            out.append(row)

    path = AUDIT / "codes.csv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["post_id"] + FIELDS)
        w.writeheader()
        w.writerows(out)

    print(f"{path}: {len(out)} row(s), {len(out)-len(unmatched)} coded, "
          f"{len(unmatched)} unclear")
    for b, d, c in unmatched:
        print(f"    UNMATCHED  {b[:28]:30} {d:11} {c}")


if __name__ == "__main__":
    main()
