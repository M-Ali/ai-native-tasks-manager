"""Codes for Engro Fertilizers' 24-month YouTube window (110 uploads, 5 Oct 2024 - 2 Sep 2026).

Same nine-dimension schema and same method as the Sarsabz read (`owned/code_window.py`): coded from
thumbnail, title and description; videos not watched end to end.

Two claim values are added here because Engro makes claims Sarsabz does not:
  - `cost_saving`      "Kam Kharch, Zyada Paidawar", "Kharcha kam aur munafa ziada"
  - `authenticity`     "DAP Bharosa Seal - chamkay toh asal, na chamkay toh naqal" + SMS verification
Both were checked against the whole Sarsabz channel before being added: Sarsabz made an authenticity
claim in 2021-22 ("Sarsabz Ki Asli Nitrophos") and a value claim in 2017 ("Khaad Muft Hi Samjho"),
neither of them inside its 24-month window. So these are abandoned Sarsabz territories, now occupied.
"""
import csv

VOCAB = {
    "content_type": ["brand_building", "tactical_promo", "product_feature", "corporate_pr", "recruitment",
                     "calendar_topical", "csr", "ugc_repost", "sponsorship"],
    "claim_primary": ["yield_increase", "product_quality", "advisory_service", "farmer_recognition",
                      "women_empowerment", "national_identity", "rewards_prizes", "brand_goodwill",
                      "cost_saving", "authenticity", "other"],
    "register": ["fear", "aspiration", "duty", "reassurance", "pride", "humour"],
    "proof_device": ["sovereign_guarantee", "ratings_awards", "testimonial", "claim_statistics", "celebrity",
                     "expert_authority", "heritage_scale", "none"],
    "who_is_in_frame": ["customer", "celebrity", "executive", "staff", "expert", "none_people"],
    "production_value": ["studio", "stock", "template_graphic", "event_photo", "ugc", "ai_generated", "unclear"],
    "audience_generation": ["genz", "geny", "genx", "boomer_plus", "mixed", "unclear"],
}
F = ["platform", "content_type", "claim_primary", "register", "proof_device", "who_is_in_frame",
     "production_value", "audience_generation", "product_focus", "ecosystem_push", "addresses_objection", "notes"]

T = {
    "agronomy_fact": ("agronomy_education", "product_feature", "advisory_service", "reassurance", "expert_authority",
                      "none_people", "template_graphic", "genx", "soil nutrition", "no", "no",
                      "nutrient explainer card, e.g. '98% of Pakistani soils are nitrogen deficient'"),
    "expert_tips": ("engro_kahanian", "product_feature", "advisory_service", "reassurance", "expert_authority",
                    "expert", "studio", "genx", "crop advisory", "no", "yes",
                    "named agronomist, crop-specific problem solved to camera"),
    "shaandar": ("shaandar_kissan", "brand_building", "farmer_recognition", "pride", "testimonial", "customer",
                 "studio", "genx", "brand", "no", "yes",
                 "named progressive farmer, long-form interview on his own farm"),
    "zabardast": ("zabardast_urea", "product_feature", "product_quality", "aspiration", "claim_statistics",
                  "celebrity", "studio", "geny", "Engro Zabardast Urea", "no", "yes",
                  "bio-active zinc coating; explicit arithmetic - one bag = ordinary urea + 2 zinc packets"),
    "zarkhez": ("zarkhez", "product_feature", "product_quality", "aspiration", "none", "none_people", "studio",
                "genx", "Engro Zarkhez Khas/Plus", "no", "no", "packshot + crop, blend and micronutrient claim"),
    "markaz": ("engro_markaz", "product_feature", "advisory_service", "reassurance", "none", "customer", "studio",
               "genx", "Engro Markaz retail + app", "yes", "yes",
               "owned retail centre: stock, advice and app in one; 'Kisaan ka Markaz e Yaqeen'"),
    "mitti_mahir": ("mitti_mahir", "product_feature", "advisory_service", "reassurance", "expert_authority",
                    "staff", "studio", "genx", "Engro Mitti Mahir soil testing", "yes", "yes",
                    "soil testing service - 'ab andaza nahi, tajzia zaroori hai' (no more guessing, test it)"),
    "jeet_baazi": ("jeet_ki_baazi", "tactical_promo", "rewards_prizes", "aspiration", "none", "none_people",
                   "studio", "genx", "Engro Urea/DAP", "yes", "no",
                   "scratch-and-SMS lucky draw: pickup truck, motorcycles, air tickets"),
    "corporate": ("corporate", "corporate_pr", "brand_goodwill", "pride", "expert_authority", "executive",
                  "studio", "unclear", "Engro corporate", "no", "no",
                  "executive or third-party authority talking head, food-security framing"),
    "ugai": ("ugai_app", "product_feature", "advisory_service", "reassurance", "expert_authority", "none_people",
             "studio", "geny", "#UgAi advisory app", "yes", "no", "AI crop-advisory app launch"),
    "musafir": ("engro_musafir", "brand_building", "brand_goodwill", "pride", "none", "customer", "studio",
                "mixed", "brand", "no", "no", "'Mitti, logon aur tajurbon ka safarnama' brand documentary strand"),
    "gupshup": ("engro_gupshup", "product_feature", "cost_saving", "reassurance", "expert_authority", "expert",
                "studio", "genx", "cost per acre", "no", "yes",
                "podcast: 'Kam Kharch, Zyada Paidawar' - saving and profit in maize and rice"),
    "bharosa": ("dap_bharosa_seal", "product_feature", "authenticity", "reassurance", "claim_statistics",
                "celebrity", "studio", "genx", "Engro DAP", "yes", "yes",
                "anti-counterfeit seal: 'chamkay toh asal, na chamkay toh naqal' + SMS verification"),
}
C = {1: "corporate", 2: "agronomy_fact", 3: "agronomy_fact", 5: "agronomy_fact",
     4: "ugai", 7: "ugai",
     6: "corporate", 8: "corporate", 9: "corporate", 10: "corporate", 11: "corporate", 12: "corporate",
     13: "corporate", 14: "corporate", 15: "corporate", 16: "corporate", 17: "corporate",
     21: "markaz", 22: "corporate", 23: "corporate",
     18: "zabardast", 39: "zabardast", 40: "zabardast", 41: "zabardast", 42: "zabardast", 46: "zabardast",
     47: "zabardast", 98: "zabardast", 102: "zabardast", 103: "zabardast",
     19: "zarkhez", 25: "zarkhez", 26: "zarkhez", 27: "zarkhez", 90: "zarkhez", 92: "zarkhez", 93: "zarkhez",
     94: "zarkhez", 95: "zarkhez",
     20: "shaandar", 43: "shaandar", 44: "shaandar", 45: "shaandar", 48: "shaandar", 58: "shaandar",
     28: "expert_tips", 29: "expert_tips", 30: "expert_tips", 31: "expert_tips", 32: "expert_tips",
     33: "expert_tips", 34: "expert_tips", 35: "expert_tips", 36: "expert_tips", 37: "expert_tips",
     49: "expert_tips", 50: "expert_tips", 53: "expert_tips", 55: "expert_tips", 56: "expert_tips",
     60: "expert_tips", 70: "expert_tips", 74: "expert_tips", 75: "expert_tips", 77: "expert_tips",
     78: "expert_tips", 79: "expert_tips", 80: "expert_tips", 81: "expert_tips", 82: "expert_tips",
     38: "gupshup", 57: "gupshup", 59: "gupshup", 61: "gupshup", 107: "gupshup", 109: "gupshup",
     51: "markaz", 52: "markaz", 54: "markaz", 97: "markaz", 99: "markaz", 100: "markaz", 101: "markaz",
     108: "markaz",
     63: "jeet_baazi", 64: "jeet_baazi", 65: "jeet_baazi", 67: "jeet_baazi", 68: "jeet_baazi", 69: "jeet_baazi",
     72: "jeet_baazi", 73: "jeet_baazi", 76: "jeet_baazi",
     83: "mitti_mahir", 84: "mitti_mahir", 85: "mitti_mahir", 86: "mitti_mahir", 87: "mitti_mahir",
     88: "musafir", 89: "musafir", 91: "musafir", 96: "musafir",
     104: "bharosa", 105: "bharosa", 110: "bharosa",
     }


def o(platform, content_type, claim, register, proof, who, prod, gen, focus, notes, eco="no", obj="no"):
    return dict(zip(F, (platform, content_type, claim, register, proof, who, prod, gen, focus, eco, obj, notes)))


O = {
    24: o("np_plus", "product_feature", "product_quality", "aspiration", "none", "none_people", "studio", "genx",
          "Engro NP+", "NP+ packshot - Engro competing in Fatima's own NP category"),
    62: o("tsp", "product_feature", "product_quality", "aspiration", "none", "none_people", "studio", "genx",
          "Engro TSP (imported)", "TSP launch: phosphate alternative to DAP, imported"),
    66: o("tsp", "product_feature", "product_quality", "aspiration", "none", "none_people", "studio", "genx",
          "Engro TSP (imported)", "TSP packshot"),
    71: o("tsp", "product_feature", "product_quality", "aspiration", "none", "none_people", "studio", "genx",
          "Engro TSP (imported)", "TSP 15s cut"),
    106: o("corporate", "calendar_topical", "national_identity", "pride", "none", "none_people", "studio", "mixed",
           "Engro corporate", "'Kisaan Zindabad, Pakistan Paindabad' - Independence framing"),
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
        out.append({"n": n, "brand": "Engro Fertilizers", "video_id": r["video_id"], "tab": r["tab"],
                    "published": r["published"], "views": int(r["views_exact"] or r["views"] or 0),
                    "length_sec": int(r["length_exact"] or r["length_sec"] or 0),
                    "title": r["title"], "url": r["url"], **rec, "coder": "claude-thumbnail+description"})
    with open("codes_window.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(out)
    print(len(out), "rows coded and validated")
