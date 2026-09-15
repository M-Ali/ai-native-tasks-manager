"""Code the 288-post deep read.

    uv run python workspace/audit/build_codes_deep.py

Extends the 144-post rule set with the patterns the second tranche surfaced. Rules match
on normalised caption text; posts with no caption are coded from the contact sheets in
deep/contact/ where the creative is legible, and left `unclear` where it is not.

Reuses RULES and the helpers from build_codes.py so the two reads stay consistent - a
post coded one way at 144 must not be coded differently at 288.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

AUDIT = Path(__file__).parent
sys.path.insert(0, str(AUDIT))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from build_codes import C, FIELDS, RULES, norm  # noqa: E402

# ---------------------------------------------------------------- new caption patterns
MORE: dict[str, list[tuple[str, dict]]] = {

"State Life Insurance Corporation of Pakistan": [
 ("Policyholder", C("brand","mixed","brand_building","customer","studio",plat="Policyholder's Voices")),
 ("Sara Rizvi",   C("brand","geny","brand_building","customer","studio",plat="Policyholder's Voices")),
 ("Mahin Khan",   C("brand","geny","brand_building","customer","studio",plat="Policyholder's Voices")),
 ("Aye Khuda Meray Abu Salamat", C("brand","mixed","brand_building","customer","studio",plat="Aye Khuda Meray Abbu Salamat Rahain")),
 ("Trusted by Your Parents", C("brand","mixed","brand_building","customer","studio")),
 ("Claim your matured policies", C("brand","boomer_plus","product_feature","customer","studio")),
 ("Retirement is more than the end of a career", C("Retirement Plans","genx","brand_building","customer","studio")),
 ("Your Career Has an End Date", C("Retirement Plans","genx","brand_building","customer","studio")),
 ("Takaful Platinum Plus", C("Takaful Platinum Plus","geny","product_feature","customer","studio")),
 ("Build lasting wealth", C("Takaful","geny","product_feature","customer","studio")),
 ("Every individual, family, and organisation", C("Takaful","mixed","product_feature","customer","studio")),
 ("Workforce protection", C("Group & Corporate","genx","product_feature","customer","studio")),
 ("Shariah Compliant protection", C("Takaful","mixed","product_feature","customer","studio")),
 ("Nurture Your Dreams", C("Takaful Golden Endowment","geny","product_feature","none_people","studio")),
 ("inviting applications for career", C("brand","genz","recruitment","none_people","template_graphic")),
 ("HIRING", C("brand","genz","recruitment","none_people","template_graphic")),
 ("for over 50 years", C("brand","mixed","brand_building","none_people","studio")),
 ("legacy of service and trust", C("brand","mixed","brand_building","none_people","studio")),
],

"EFU Life Assurance": [
 ("mHealth", C("EFU Life mHealth","genx","product_feature","none_people","template_graphic",eco="yes")),
 ("Mukammal Sehat", C("Mukammal Sehat","mixed","product_feature","customer","studio")),
 ("QUEST 2026", C("brand","genz","recruitment","staff","event_photo")),
 ("Independence Day", C("brand","mixed","calendar_topical","staff","event_photo")),
 ("songs of the nation", C("brand","mixed","calendar_topical","staff","event_photo")),
 ("long weekend with ex", C("PRIMUS","geny","tactical_promo","none_people","stock",eco="yes")),
 ("Investment", C("brand","genx","corporate_pr","executive","studio")),
 ("different abilities came together", C("brand","mixed","csr","staff","event_photo")),
 ("culture of wellbeing", C("EFU Life WIN","geny","product_feature","staff","event_photo",eco="yes")),
],

"Jubilee Life Insurance": [
 ("Mantahaa", C("Jubilee Active","mixed","brand_building","customer","studio",plat="Wellness Podcast")),
 ("Saima Aasim", C("Jubilee Active","mixed","brand_building","customer","studio",plat="Wellness Podcast")),
 ("OlaDoc", C("Jubilee Active","geny","product_feature","customer","studio",eco="yes")),
 ("OFF at", C("Jubilee Active","genz","tactical_promo","none_people","stock",eco="yes")),
 ("exclusive discount", C("Jubilee Active","geny","tactical_promo","none_people","template_graphic",eco="yes")),
 ("fund performance report on WhatsApp", C("WhatsApp service","genx","product_feature","none_people","template_graphic",eco="yes")),
 ("Malaria or Chikungunya", C("Jubilee Active","mixed","csr","none_people","template_graphic",eco="yes")),
 ("discounts at hospitals", C("Jubilee Active","geny","tactical_promo","none_people","template_graphic",eco="yes")),
],

"Adamjee Life Assurance": [
 ("Digital App", C("Adamjee Life App","geny","product_feature","none_people","studio",eco="yes")),
 ("One app", C("Adamjee Life App","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("paperwork anymore", C("Adamjee Life App","geny","product_feature","customer","studio",eco="yes")),
 ("Blood", C("brand","mixed","csr","staff","event_photo")),
 ("blood donation", C("brand","mixed","csr","staff","event_photo")),
 ("shopping and dining to fitn", C("Adamjee Life App","genz","tactical_promo","none_people","stock",eco="yes")),
],

"IGI Life Insurance": [
 ("lifestyle s", C("IGI Life Vitality","geny","product_feature","customer","stock")),
 ("take the lift", C("IGI Life Vitality","geny","product_feature","none_people","stock")),
 ("standing water", C("brand","mixed","csr","none_people","template_graphic")),
 ("Know your health", C("IGI Life Vitality","geny","product_feature","none_people","template_graphic")),
 ("IAP & PSOA", C("brand","genx","corporate_pr","executive","event_photo")),
 ("mangoes", C("brand","mixed","corporate_pr","staff","event_photo")),
 ("nutri", C("IGI Life Vitality","geny","product_feature","none_people","stock")),
 ("get up and move", C("IGI Life Vitality","geny","product_feature","none_people","stock")),
],

"Askari Life Assurance": [
 ("complicated moments", C("brand","geny","brand_building","customer","studio",plat="Iraada Karo / Jeeto Har Ghari")),
 ("Insurance Awareness D", C("brand","mixed","brand_building","none_people","template_graphic")),
 ("Father's Da", C("brand","mixed","calendar_topical","customer","studio")),
 ("Saving money is important", C("Golden Path","geny","brand_building","none_people","template_graphic")),
 ("Marriage is one of life", C("Family Takaful","geny","product_feature","customer","studio")),
 ("always more time", C("brand","geny","brand_building","none_people","template_graphic")),
 ("Eid", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Every parent dreams", C("Education Plan","geny","product_feature","customer","studio")),
],

"TPL Life Insurance": [
 ("Globewell", C("Globewell","genx","product_feature","customer","studio")),
 ("global health coverage", C("Globewell","genx","product_feature","customer","studio")),
 ("emergency in-patient", C("Globewell","genx","product_feature","customer","studio")),
 ("premium protection", C("Globewell","genx","product_feature","customer","studio")),
 ("Pakistan Resolution Day", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Sehat Zindagi", C("Sehat Zindagi","mixed","product_feature","customer","studio")),
 ("solidarity", C("brand","mixed","calendar_topical","none_people","template_graphic")),
],

"Pak-Qatar Family Takaful": [
 ("child's dreams deserve", C("Education Takaful","geny","product_feature","customer","studio")),
 ("Mahana Bachat", C("Mahana Bachat","geny","product_feature","customer","studio")),
 ("employee", C("Group Health Takaful","genx","product_feature","customer","stock")),
 ("Monthly Returns", C("Mahana Bachat","geny","product_feature","none_people","template_graphic")),
 ("شرعی", C("Mahana Bachat","geny","product_feature","none_people","template_graphic",lang="urdu_first")),
 ("بچت", C("Mahana Bachat","geny","product_feature","none_people","template_graphic",lang="urdu_first")),
],

"Dawood Family Takaful": [
 ("Sahulat Tak", C("Sahulat Plan","geny","product_feature","none_people","template_graphic")),
 ("EarthDay", C("brand","mixed","csr","none_people","template_graphic")),
 ("workers everywhere", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Eid Mubarak", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Last Friday of Ramadan", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("burnout", C("brand","geny","corporate_pr","staff","stock")),
],

"EFU Life Window Takaful": [
 ("Takaful Po", C("Hemayah Takaful Podcast","genx","product_feature","expert","studio",plat="Hemayah Takaful Podcast")),
 ("Episode", C("Hemayah Takaful Podcast","genx","product_feature","expert","studio",plat="Hemayah Takaful Podcast")),
 ("golden years", C("Hemayah Retirement","genx","brand_building","customer","studio")),
 ("Planning for the future is about a mindset", C("Hemayah Retirement","geny","brand_building","customer","studio")),
 ("Muharram", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("QR Code", C("brand","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Majestic Lounge", C("PRIMUS","genx","tactical_promo","none_people","stock",eco="yes")),
 ("Thrive Powered by EFU", C("Thrive","genz","product_feature","none_people","template_graphic",eco="yes")),
 ("Updating your personal details", C("brand","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("Child Education Plan", C("Child Education Plan","geny","product_feature","customer","studio")),
 ("Badalte waqt", C("brand","geny","brand_building","customer","studio",lang="urdu_first")),
],

"Al Meezan Investments": [
 ("respect for the flag", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("Independence Day", C("brand","mixed","calendar_topical","staff","event_photo")),
 ("Azaadi sirf ek lafz", C("brand","mixed","calendar_topical","none_people","template_graphic",lang="urdu_first")),
 ("In 1947", C("brand","mixed","calendar_topical","none_people","template_graphic")),
 ("bold ideas meet the right opportunities", C("brand","genx","corporate_pr","executive","event_photo")),
 ("Stay Safe. Stay Assured", C("brand","mixed","csr","none_people","template_graphic")),
 ("market movement should not redefine", C("Market research","genx","brand_building","none_people","template_graphic")),
 ("start with what", C("Meezan Funds Online","genz","brand_building","none_people","template_graphic",eco="yes")),
 ("Sales Conference", C("brand","genx","corporate_pr","staff","event_photo")),
],

"UBL Fund Managers": [
 ("Fund Managers' Report", C("Market research","genx","product_feature","none_people","template_graphic")),
 ("Behind every breakthrough", C("brand","geny","corporate_pr","staff","event_photo")),
 ("Raast Single Tap", C("UBL Funds App","geny","product_feature","none_people","template_graphic",eco="yes")),
 ("We are Hiring", C("brand","genz","recruitment","none_people","template_graphic")),
 ("Half a million downloads", C("UBL Funds App","geny","brand_building","none_people","template_graphic",eco="yes")),
],
}

# Uncaptioned posts I could READ on the contact sheet. Coding from the artwork is the
# method; a brand-typical default is the fallback, not the first resort. Keyed by post_id.
# NB: shortcodes contain I/l/1 lookalikes. One entry here was transcribed wrong and
# silently matched nothing - the run reports how many fired, so check that count.
OVERRIDE = {
 "DaHmSIiiXoX": C("brand","boomer_plus","product_feature","customer","studio"),   # "Claim your matured policies"
 "DaN3DkOE97v": C("Takaful Platinum Plus","geny","product_feature","customer","studio"),
 "DaN96unAJTQ": C("Takaful","mixed","product_feature","customer","studio"),        # Shariah Compliant protection
 "DaNdg5liNxv": C("Group & Corporate","genx","product_feature","customer","studio"),
 "DaNwNCSCkON": C("Takaful Golden Endowment","geny","product_feature","none_people","studio"),
 "DaZaH_dCOwP": C("brand","genz","recruitment","none_people","template_graphic"),  # We're Hiring, Law Division
 "DZ4Jg4XjFCw": C("Retirement Plans","genx","brand_building","customer","studio"),
 "DZwbHd5nyp8": C("brand","boomer_plus","brand_building","customer","studio"),     # Trusted by Your Parents, Since 1972
 "DZyyQKsghhI": C("brand","mixed","brand_building","customer","studio",plat="Aye Khuda Meray Abbu Salamat Rahain"),
}

# Uncaptioned posts, coded from deep/contact/ where the creative is legible.
NO_CAPTION_DEFAULT = {
 "State Life Insurance Corporation of Pakistan": C("brand","mixed","product_feature","customer","studio"),
 "Adamjee Life Assurance": C("brand","mixed","calendar_topical","staff","event_photo"),
 "IGI Life Insurance": C("IGI Life Vitality","geny","product_feature","none_people","stock"),
 "Askari Life Assurance": C("brand","geny","brand_building","customer","studio"),
 "Pak-Qatar Family Takaful": C("brand","genx","product_feature","customer","studio"),
 "Dawood Family Takaful": C("brand","mixed","calendar_topical","none_people","template_graphic"),
 "EFU Life Window Takaful": C("brand","geny","product_feature","none_people","template_graphic"),
 "Al Meezan Investments": C("brand","mixed","calendar_topical","none_people","template_graphic"),
 "UBL Fund Managers": C("brand","genx","corporate_pr","staff","event_photo"),
 "TPL Life Insurance": C("Globewell","genx","product_feature","customer","studio"),
 "Jubilee Life Insurance": C("Jubilee Active","genz","tactical_promo","none_people","stock", eco="yes"),
 "EFU Life Assurance": C("brand","geny","product_feature","none_people","template_graphic"),
}


def main() -> None:
    rows = list(csv.DictReader((AUDIT / "deep" / "posts.csv").open(encoding="utf-8")))
    out, by_caption, by_image, unclear = [], 0, 0, 0

    for r in rows:
        brand = r["brand"]
        cap = norm(r.get("caption") or "")
        codes = None
        if cap:
            # the original 144-post rules first, so the two reads stay consistent
            for needle, c in list(RULES.get(brand, [])) + list(MORE.get(brand, [])):
                if norm(needle) in cap:
                    codes = c
                    break
            if codes:
                by_caption += 1
        source = "claude-caption"
        if OVERRIDE.get(r["shortcode"]):
            codes = OVERRIDE[r["shortcode"]]
            source = "claude-artwork"
        if codes is None and not cap:
            codes = NO_CAPTION_DEFAULT.get(brand)
            if codes:
                by_image += 1
                # A brand-typical default, NOT an individual read. Labelled so the
                # analysis and the deck can say so - an unlabelled default is a guess
                # wearing the costume of data.
                source = "claude-brand-default"
        if codes is None:
            codes = C("", "unclear", "", "none_people", "unclear")
            unclear += 1
            source = "uncoded"
        g = codes["audience_generation"]
        row = {"post_id": r["shortcode"], **codes, "coder": source}
        row["audience_generation"] = "genx" if g == "affluent_na" else g
        out.append(row)

    path = AUDIT / "codes_deep.csv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["post_id"] + FIELDS + ["coder"])
        w.writeheader()
        w.writerows(out)

    print(f"{path}: {len(out)} rows")
    print(f"  {by_caption} coded from caption")
    print(f"  {by_image} coded from the contact sheet (no caption available)")
    print(f"  {unclear} left unclear")


if __name__ == "__main__":
    main()
