"""Code every Salam Kissan / Kissan Day upload on @Sarsabz (2019-2025) -> archive_codes.csv + archive_findings.md.

Coded after reading all seven contact sheets (contact/sheet_01-07.jpg) with titles and descriptions.
Videos were not watched end to end.

What the codes answer, because they are the brief's Case 1 asks (p12):
- who speaks          does the farmer have a voice, or is he the subject of someone else's tribute?
- audience            is the work addressed to farmers, to urban Pakistan, or to both?
- in_december         does the work live beyond the day? (1 Dec - 15 Jan counts as "the day", event recaps included)
- format              what the platform is actually made of, year by year
"""
import csv
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).parent

# format, speaker, audience, note
C = {}
def put(ns, fmt, speaker, audience, note):
    for n in ns:
        C[n] = (fmt, speaker, audience, note)

put([1, 2], "celebrity_testimonial", "celebrity", "urban", "Jami (filmmaker), Ali Noor (musician) on why farmers matter")
put([3, 4, 6], "event", "participants", "mixed", "Canvas Wall: art teams paint tributes at Pak Arab plant, Multan")
put([5], "event", "celebrity", "mixed", "Canvas Wall live with host Sophiya Anjam")
put([7], "anthem_film", "none_spoken", "mass", "first anthem: 'A Tribute to Farmers', Ali Noor vocals, Jami directs")
put([8], "pr_media", "executive", "urban", "GTV news segment, Fatima head of technical")
put([9, 10, 11, 14], "celebrity_testimonial", "celebrity", "urban", "influencer / journalist 'celebrating Kissan Day on 18th Dec'")
put([12, 13, 17], "farmer_testimonial", "farmer", "mixed", "farmer or farmer leader endorses the day (Nizamani, Rabia Sultan, Azra Mehmood Sheikh)")
put([15, 16], "celebrity_testimonial", "celebrity", "urban", "cricketer / actor 'celebrating Kissan Day'")
put([18], "politician_testimonial", "politician", "mixed", "Food Minister Punjab")
put([19], "behind_the_scenes", "crew", "mass", "making of the anthem")
put([20], "event", "participants", "mixed", "Kissan Day 2019 event highlights")
put([21, 22], "calendar_other", "farmer", "mixed", "Women's Day: female kissan speak (Mar 2020)")
put([23], "calendar_other", "none_spoken", "urban", "'Mulk Ke Heroes' Eid film, urban child (May 2020)")
put([24], "politician_testimonial", "politician", "mixed", "Speaker National Assembly at wheat campaign (Oct 2020)")
put([25], "award_pr", "none_spoken", "industry", "Best Digital Campaign, Pakistan Digital Awards 2020")
put([26], "urban_short", "none_spoken", "urban", "'Let's pledge not to waste food' - urban boy eating an apple")
put([27], "urban_short", "none_spoken", "urban", "'Thank you, Kissan!' - basket of produce in an urban home")
put([28], "urban_short", "none_spoken", "urban", "'Time to appreciate the unsung heroes' - urban girl at a table")
put([29, 30, 32, 34, 35, 38, 39, 42, 43, 44], "celebrity_testimonial", "celebrity", "urban",
    "celebrity 'celebrating Kissan Day on 18th Dec' (Afridi, Wasim Akram, Ahsan Khan, Muniba Mazari, Reema Khan...)")
put([31, 33, 36, 37, 41], "politician_testimonial", "politician", "mixed",
    "minister thanks farmers and Sarsabz (Punjab, KPK, federal food security, climate change)")
put([40], "partner_testimonial", "executive", "industry", "CMEC managing director (Chinese partner)")
put([45], "event", "participants", "mixed", "Kissan Day 2020, Governor House Lahore")
put([46], "pr_media", "executive", "mixed", "Chairman Fatima Group message (Jan 2021)")
put([47], "politician_testimonial", "politician", "mixed", "Governor Punjab speech at Kissan Day event (Jan 2021)")
put([48], "farmer_testimonial", "farmer", "mixed", "President Pakistan Kissan Ittehad speech (Jan 2021)")
put([49], "product_testimonial", "farmer", "farmer", "rice testimonial 'Kissan Tera Ehsaan, Sarsabz Pakistan' (Jul 2021)")
put([50], "anthem_film", "none_spoken", "mass", "Salam Kissan 2021 anthem - 13.1M")
put([51], "event", "participants", "mixed", "Kissan Day at Pakistan Pavilion, Expo 2020 Dubai")
put([52], "anthem_film", "none_spoken", "mass", "Salam Kissan 2022 anthem - 6.5M")
put([53], "urban_short", "none_spoken", "urban", "'How would you feel if this was your reality?' - urban woman, empty plate")
put([54], "urban_short", "none_spoken", "urban", "'Will you continue to ignore the hardwork of those who provide for you?'")
put([55], "event", "participants", "mixed", "Kissan Day 2022, Pak Arab plant Multan")
put([56, 57, 58, 59, 60, 61], "product_testimonial", "farmer", "farmer",
    "named wheat farmer by district, NP + CAN yield - product films in the Kissan Day window (Nov-Dec 2023)")
put([62], "anthem_film", "none_spoken", "mass", "Salam Kissan 2023 anthem - 22.6M, the peak")
put([63], "corporate_pr", "none_spoken", "industry", "UNDP SDG partnership film (Oct 2024)")
put([64, 65], "farmer_salute_short", "farmer", "farmer", "15s farmer salute to camera - the whole 2024 Kissan Day output")
put([66], "behind_the_scenes", "crew", "industry", "Salam Kissan journey (Jan 2025)")
put([67], "event", "politician", "mixed", "Kissan Day 2024, Islamabad - Federal Minister Rana Tanveer (Jan 2025)")
put([68], "csr", "children", "urban", "SOS Children's Village art competition (May 2025)")
put([69, 72, 73, 74, 75], "animated_short", "none_spoken", "mass", "3D/AI animated farmer, 'Celebrating Kissan Day on 18th December'")
put([70], "urban_short", "celebrity", "urban", "Salam Kissan 2025: actress at a dining table - 5.4M")
put([71], "farmer_mosaic", "farmer", "mixed", "mosaic of farmer faces - 2.2M")
put([76], "urban_short", "celebrity", "urban", "Ali Rehman, food blogger: 'Behind every meal is a farmer's hard work' - 4M")


def in_window(date: str) -> bool:
    """'The day' = 1 December to 15 January (covers pre-launch and event recaps)."""
    mmdd = date[5:10] if len(date) >= 10 and "x" not in date else date[5:7] + "-15"
    return mmdd >= "12-01" or mmdd <= "01-15"


rows = list(csv.DictReader((HERE / "sk_index.csv").open(encoding="utf-8")))
missing = [r["n"] for r in rows if int(r["n"]) not in C]
if missing:
    raise SystemExit(f"uncoded uploads: {missing}")
out = []
for r in rows:
    fmt, speaker, audience, note = C[int(r["n"])]
    out.append({"n": r["n"], "video_id": r["video_id"], "published": r["published"], "year": r["published"][:4],
                "views": int(float(r["views"] or 0)), "title": r["title"], "format": fmt, "speaker": speaker,
                "audience": audience, "in_window": "yes" if in_window(r["published"]) else "no", "note": note})
with (HERE / "archive_codes.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(out[0])); w.writeheader(); w.writerows(out)

N, V = len(out), sum(r["views"] for r in out)
L = [f"# Salam Kissan archive - coded read (2019-2025)", "",
     f"**{N} uploads on @Sarsabz, {V / 1e6:.1f}M views.** Coded from all seven contact sheets plus titles and descriptions; "
     "not watched end to end. Views on a brand channel are mostly bought - they show where the brand put weight.", ""]

def table(key, title):
    cnt, vw = Counter(), Counter()
    for r in out:
        cnt[r[key]] += 1; vw[r[key]] += r["views"]
    L.extend([f"## {title}", "", f"| {key} | uploads | views (M) | share of views |", "|---|---:|---:|---:|"])
    for k in sorted(cnt, key=lambda k: -vw[k]):
        L.append(f"| {k} | {cnt[k]} | {vw[k] / 1e6:.1f} | {100 * vw[k] / V:.0f}% |")
    L.append("")

table("format", "What the platform is made of")
table("speaker", "Who speaks")
table("audience", "Who it addresses")

L.extend(["## By year", "", "| year | uploads | views (M) | anthem film | biggest upload |", "|---|---:|---:|---|---|"])
by = defaultdict(list)
for r in out:
    by[r["year"]].append(r)
for y in sorted(by):
    rs = by[y]; top = max(rs, key=lambda r: r["views"])
    anthem = next((f"{r['views'] / 1e6:.1f}M" for r in rs if r["format"] == "anthem_film"), "none")
    L.append(f"| {y} | {len(rs)} | {sum(r['views'] for r in rs) / 1e6:.1f} | {anthem} | {top['title'][:48]} ({top['views'] / 1e6:.1f}M) |")
L.append("")

inw = [r for r in out if r["in_window"] == "yes"]
L.extend(["## Beyond the day?", "",
          f"**{len(inw)} of {N} uploads ({100 * len(inw) / N:.0f}%) and {100 * sum(r['views'] for r in inw) / V:.0f}% of views fall between "
          "1 December and 15 January.** Outside that window: " +
          "; ".join(f"#{r['n']} {r['note'][:50]}" for r in out if r["in_window"] == "no") + ".", ""])

farmer_voice = [r for r in out if r["speaker"] == "farmer" and r["format"] not in ("product_testimonial",)]
L.extend(["## Does the farmer speak?", "",
          f"Uploads where a farmer speaks for himself or herself, outside product testimonials: **{len(farmer_voice)} of {N}**, "
          f"{sum(r['views'] for r in farmer_voice) / 1e6:.1f}M views. Celebrities and politicians speak in "
          f"{sum(1 for r in out if r['speaker'] in ('celebrity', 'politician'))}.", ""])
(HERE / "archive_findings.md").write_text("\n".join(L) + "\n", encoding="utf-8")
print("\n".join(L))
