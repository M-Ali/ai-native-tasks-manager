"""Channel inventory: every Sarsabz YouTube video/short with approximate date, views, length,
and a rule-based campaign-platform tag from the TITLE only (a sorting aid, not a code)."""
import csv, json, re
from datetime import datetime, timezone

PLATFORMS = [  # first match wins; order matters
    ("salam_kissan", r"salam ?kiss?an|kissan day|kisan day"),
    ("ki_jeet_competition", r"ki jeet|paidawaari muqabla|gandum.*muqabla|yield competition|convention"),
    ("yield_10pct_product", r"10 ?%|10 feesad|izaafi|izafi|nitrophos aur can|np aur can"),
    ("nitrophos_scheme_promo", r"scheme|golden ticket|lucky draw|inaam|prize"),
    ("np_product_other", r"nitrophos|\bnp\b|\bcan\b|calcium ammonium|dap|urea"),
    ("tabeer_women", r"tabeer|women|beti|iwd"),
    ("kahani_film", r"kahani|short film"),
    ("ramzan_religious", r"ramzan|ramadan|dua|naat|eid|hajj|muharram|qaseed"),
    ("sport_sponsorship", r"sultan|psl|kabaddi|kabbadi|tent pegg|cricket"),
    ("national_calendar", r"independence|14 ?august|23 ?march|pakistan day|dil se dekho|defence day"),
    ("advisory_agronomy", r"advis|mashwara|kaasht|kasht|sowing|spray|irrigat|weather|mausam|tips|training|soil|mitti|beej|seed|crop|fasal|wheat|gandum|cotton|kapas|rice|dhaan|maize|makai|sugarcane|gana|app"),
    ("testimonial", r"testimonial|farmer story|kisan ki zubani"),
    ("corporate_award", r"award|pda|drum|dragons|ceo|mou|launch|ceremony|webinar"),
]
def tag(title):
    t = (title or "").lower()
    for name, rx in PLATFORMS:
        if re.search(rx, t): return name
    return "other"

rows = []
for tab in ("videos", "shorts"):
    d = json.load(open(f"channel_{tab}.json", encoding="utf-8"))
    for e in d.get("entries", []):
        ts = e.get("timestamp")
        rows.append({
            "video_id": e.get("id"), "tab": tab,
            "approx_date": datetime.fromtimestamp(ts, timezone.utc).date().isoformat() if ts else "",
            "views": e.get("view_count") or "", "length_sec": e.get("duration") or "",
            "title": e.get("title"), "platform_tag": tag(e.get("title")),
            "url": f"https://www.youtube.com/watch?v={e.get('id')}" if tab == "videos" else f"https://www.youtube.com/shorts/{e.get('id')}",
        })
with open("channel_inventory.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
print(len(rows), "rows")
