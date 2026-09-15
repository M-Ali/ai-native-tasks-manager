"""Codes for the 24-month owned-channel window (90 YouTube videos/shorts, 15 Sep 2024 - 14 Sep 2026).

Coded by reading every thumbnail (contact/sheet_01-08.jpg) plus title and description. Videos were
NOT watched end to end: register and proof are coded from what the thumbnail, title and
description commit to. Where that is a limit it is noted.

Schema: category-creative-scan nine dimensions. Deviations, stated rather than hidden:
  - claim vocabulary rewritten for fertilizer (the skill requires this per category)
  - content_type adds `sponsorship` - 26 Multan Sultans items would otherwise be forced into
    brand_building or corporate_pr and distort the ratio the scan exists to measure.
"""
import csv

VOCAB = {
    "content_type": ["brand_building", "tactical_promo", "product_feature", "corporate_pr", "recruitment",
                     "calendar_topical", "csr", "ugc_repost", "sponsorship"],
    "claim_primary": ["yield_increase", "product_quality", "advisory_service", "farmer_recognition",
                      "women_empowerment", "national_identity", "rewards_prizes", "brand_goodwill", "other"],
    "register": ["fear", "aspiration", "duty", "reassurance", "pride", "humour"],
    "proof_device": ["sovereign_guarantee", "ratings_awards", "testimonial", "claim_statistics", "celebrity",
                     "expert_authority", "heritage_scale", "none"],
    "who_is_in_frame": ["customer", "celebrity", "executive", "staff", "expert", "none_people"],
    "production_value": ["studio", "stock", "template_graphic", "event_photo", "ugc", "ai_generated", "unclear"],
    "audience_generation": ["genz", "geny", "genx", "boomer_plus", "mixed", "unclear"],
}
F = ["platform", "content_type", "claim_primary", "register", "proof_device", "who_is_in_frame",
     "production_value", "audience_generation", "product_focus", "ecosystem_push", "addresses_objection", "notes"]

T = {  # templates for recurring executions
    "tabeer": ("sarsabz_tabeer", "csr", "women_empowerment", "pride", "none", "customer", "event_photo", "mixed",
               "Sarsabz Tabeer", "no", "no", "women-farmer skills training"),
    "yield10": ("yield_10pct", "product_feature", "yield_increase", "aspiration", "none", "customer", "studio", "geny",
                "NP+CAN", "no", "no",
                "'10 feesad' badge is the claim, not proof; same walking-farmer template across crops"),
    "jeet": ("sarsabz_ki_jeet", "product_feature", "yield_increase", "pride", "ratings_awards", "customer", "studio",
             "genx", "NP+CAN", "no", "yes",
             "named winner + govt competition position + maunds/acre; '34 of 40 districts'"),
    "sk_anim": ("salam_kissan", "calendar_topical", "farmer_recognition", "pride", "none", "customer", "ai_generated",
                "mixed", "brand", "no", "no",
                "3D/AI-style animated farmer, 'Celebrating Kissan Day on 18th December'"),
    "sultan": ("multan_sultans", "sponsorship", "brand_goodwill", "humour", "celebrity", "celebrity", "studio", "genz",
               "brand", "no", "no", "cricketer entertainment episode; no product or farmer"),
    "dsd": ("dil_se_dekho", "tactical_promo", "national_identity", "pride", "none", "none_people", "template_graphic",
            "genz", "brand", "no", "no", "UGC video contest: iPhone, bike, Skardu tour prizes"),
    "dsd_win": ("dil_se_dekho", "tactical_promo", "national_identity", "pride", "none", "customer", "template_graphic",
                "genz", "brand", "no", "no", "weekly winners card"),
    "dss": ("dil_se_sarsabz", "tactical_promo", "national_identity", "pride", "none", "none_people",
            "template_graphic", "genz", "brand", "no", "no", "milli naghma TikTok singing contest 2025"),
    "scheme": ("nitrophos_scheme", "tactical_promo", "rewards_prizes", "aspiration", "none", "none_people", "studio",
               "genx", "Sarsabz Nitrophos", "yes", "no",
               "scratch-code SMS/app lucky draw: Umrah ticket, motorcycle, phone"),
    "app": ("sarsabz_app", "product_feature", "advisory_service", "reassurance", "testimonial", "customer",
            "event_photo", "geny", "Sarsabz Pakistan App", "yes", "no", "farmer app testimonial at branded booth"),
}
C = {1: "tabeer", 4: "tabeer", 13: "tabeer", 26: "tabeer", 31: "tabeer", 89: "tabeer",
     3: "yield10", 33: "yield10", 34: "yield10", 35: "yield10", 36: "yield10", 65: "yield10", 66: "yield10",
     69: "yield10", 70: "yield10",
     38: "jeet", 39: "jeet", 40: "jeet", 41: "jeet",
     43: "sk_anim", 46: "sk_anim", 47: "sk_anim", 48: "sk_anim", 49: "sk_anim",
     15: "sultan", 16: "sultan", 17: "sultan", 18: "sultan", 19: "sultan", 20: "sultan", 21: "sultan",
     53: "sultan", 54: "sultan", 56: "sultan", 57: "sultan", 58: "sultan", 59: "sultan", 61: "sultan",
     62: "sultan", 63: "sultan", 64: "sultan",
     71: "dsd", 72: "dsd", 74: "dsd", 75: "dsd", 76: "dsd", 77: "dsd", 79: "dsd", 81: "dsd", 84: "dsd",
     80: "dsd_win", 86: "dsd_win", 87: "dsd_win", 88: "dsd_win",
     27: "dss", 28: "dss", 85: "scheme", 78: "scheme", 37: "app", 42: "app"}


def o(platform, content_type, claim, register, proof, who, prod, gen, focus, notes, eco="no", obj="no"):
    return dict(zip(F, (platform, content_type, claim, register, proof, who, prod, gen, focus, eco, obj, notes)))


O = {  # one-off executions
    2: o("undp_sdg", "corporate_pr", "brand_goodwill", "aspiration", "expert_authority", "customer", "studio", "mixed",
         "Fatima Fertilizer corporate", "UNDP SDG partnership film; rural girl lead; 13.7M views"),
    5: o("undp_sdg", "corporate_pr", "brand_goodwill", "pride", "expert_authority", "executive", "event_photo",
         "unclear", "Fatima Fertilizer corporate", "SDG impact report launch"),
    6: o("salam_kissan", "calendar_topical", "farmer_recognition", "pride", "none", "customer", "studio", "genx",
         "brand", "15s hand-on-heart salute"),
    7: o("salam_kissan", "calendar_topical", "farmer_recognition", "pride", "none", "customer", "studio", "genx",
         "brand", "15s Sindhi farmer salute"),
    8: o("corporate", "corporate_pr", "brand_goodwill", "reassurance", "expert_authority", "executive", "studio",
         "unclear", "Fatima Group", "51-minute group podcast"),
    9: o("salam_kissan", "corporate_pr", "farmer_recognition", "pride", "heritage_scale", "staff", "studio", "unclear",
         "brand", "Salam Kissan journey / behind the scenes"),
    10: o("salam_kissan", "corporate_pr", "farmer_recognition", "pride", "ratings_awards", "executive", "event_photo",
          "unclear", "brand", "Kissan Day 2024 event, executives with award"),
    11: o("yield_10pct", "product_feature", "yield_increase", "pride", "none", "customer", "studio", "geny",
          "CAN (Sarsabz + Pak Arab)", "'Gandum ki izafi paidawar' - no 10% figure in description"),
    12: o("womens_day", "calendar_topical", "women_empowerment", "pride", "none", "customer", "studio", "mixed",
          "Fatima Fertilizer", "'Iss Mitti Ki Beti' - signed Fatima Fertilizer, not Sarsabz"),
    14: o("ramzan", "calendar_topical", "brand_goodwill", "reassurance", "celebrity", "celebrity", "studio", "mixed",
          "brand", "Atif Aslam, dir. Asim Raza; 25.2M views"),
    22: o("pak_arab_can", "product_feature", "product_quality", "reassurance", "none", "none_people", "studio",
          "unclear", "Pak Arab CAN", "'works with less water and moisture'; bag packshot"),
    23: o("multan_sultans", "sponsorship", "brand_goodwill", "pride", "celebrity", "celebrity", "event_photo", "genz",
          "brand", "meet & greet 2025"),
    24: o("csr_other", "csr", "brand_goodwill", "pride", "none", "customer", "event_photo", "genz",
          "Fatima Fertilizer", "SOS Children's Village art competition; children not customers"),
    25: o("multan_sultans", "sponsorship", "rewards_prizes", "pride", "none", "customer", "event_photo", "geny",
          "brand", "Golden Ticket 2025 winners"),
    29: o("dil_se_sarsabz", "tactical_promo", "national_identity", "pride", "none", "customer", "studio", "genz",
          "brand", "young woman singing, TikTok contest"),
    30: o("dil_se_sarsabz", "tactical_promo", "national_identity", "pride", "celebrity", "celebrity", "ugc", "geny",
          "brand", "musician invites entries; identity of talent not verified"),
    32: o("dil_se_sarsabz", "tactical_promo", "national_identity", "pride", "none", "customer", "event_photo", "genz",
          "brand", "contest finale event"),
    44: o("salam_kissan", "calendar_topical", "farmer_recognition", "duty", "celebrity", "celebrity", "studio", "geny",
          "brand", "urban woman at dining table - 'behind every meal'; 5.4M"),
    45: o("salam_kissan", "calendar_topical", "farmer_recognition", "pride", "none", "customer", "ugc", "mixed",
          "brand", "mosaic of farmer faces; 2.2M"),
    50: o("salam_kissan", "calendar_topical", "farmer_recognition", "duty", "celebrity", "celebrity", "studio", "geny",
          "brand", "Ali Rehman, food blogger; 4.0M; urban-audience execution"),
    51: o("agrimart", "product_feature", "advisory_service", "reassurance", "none", "staff", "studio", "genx",
          "Sarsabz AgriMart", "owned retail: fixed prices + expert advice; 1,105 views", eco="yes"),
    52: o("sarsabz_tabeer", "calendar_topical", "women_empowerment", "pride", "none", "customer", "studio", "mixed",
          "Sarsabz Tabeer", "Seeds of Change, Women's Day 2026; 8.3M"),
    55: o("multan_sultans", "sponsorship", "brand_goodwill", "pride", "celebrity", "celebrity", "studio", "genz",
          "brand", "'Jeet Ka Naam, Sarsabz Ke Sultan'; 15s, 3.2M"),
    60: o("multan_sultans", "sponsorship", "brand_goodwill", "pride", "celebrity", "celebrity", "event_photo", "genz",
          "brand", "meet & greet 2026"),
    67: o("multan_sultans", "sponsorship", "rewards_prizes", "aspiration", "none", "customer", "studio", "geny",
          "brand", "Golden Ticket 2026 winners' trip"),
    68: o("yield_10pct", "product_feature", "yield_increase", "pride", "testimonial", "customer", "studio", "genx",
          "NP+CAN", "Rameez Akbar, Shikarpur: 63 maunds/acre rice; the only 10% film with a named farmer and number",
          obj="yes"),
    73: o("dil_se_dekho", "tactical_promo", "national_identity", "pride", "none", "customer", "stock", "genz", "brand",
          "hand holding phone at Minar-e-Pakistan"),
    82: o("dil_se_dekho", "tactical_promo", "national_identity", "pride", "none", "customer", "ai_generated", "genz",
          "brand", "AI-style young woman photographing mosque"),
    83: o("dil_se_dekho", "tactical_promo", "national_identity", "pride", "none", "customer", "ai_generated", "genz",
          "brand", "AI-style woman in hijab photographing mosque"),
    90: o("dil_se_dekho", "calendar_topical", "national_identity", "pride", "none", "staff", "event_photo", "unclear",
          "brand", "drone '79' human formation, Independence Day"),
}

if __name__ == "__main__":
    rows = list(csv.DictReader(open("window_index.csv", encoding="utf-8")))
    out = []
    for r in rows:
        n = int(r["n"])
        if n in C:
            rec = dict(zip(F, T[C[n]]))
        elif n in O:
            rec = O[n]
        else:
            raise SystemExit(f"uncoded row #{n}: {r['title']}")
        for k, allowed in VOCAB.items():
            if rec[k] not in allowed:
                raise SystemExit(f"#{n} {k}={rec[k]!r} not in vocabulary")
        out.append({"n": n, "video_id": r["video_id"], "tab": r["tab"], "published": r["published"],
                    "views": int(r["views_exact"] or r["views"] or 0),
                    "length_sec": int(r["length_exact"] or r["length_sec"] or 0),
                    "title": r["title"], "url": r["url"], **rec, "coder": "claude-thumbnail+description"})
    with open("codes_window.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(out)
    print(len(out), "rows coded and validated")
