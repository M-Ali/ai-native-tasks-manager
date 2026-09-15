"""Full metadata for non-Sarsabz uploads found by search.py whose channel is a brand, institution or critic.
Channel list picked by reading search_raw.csv titles; Indian 'Kisan Diwas' (23 Dec) status channels excluded. -> others.csv"""
import csv, time, sys
from yt_dlp import YoutubeDL
sys.stdout.reconfigure(encoding="utf-8")
CH = {"Syngenta Pakistan", "JPL Pakistan", "Barket Fertilizers", "Rizq Foods", "LIMS COE", "DIGITAL KISAN Pakistan",
      "Neo News", "BaKhabar Kissan ", "MANG SPECIAL", "Soby Agro Machinery", "Green Corporate Livestock Initiative",
      "Express News", "SAMAA TV", "GNN", "Aaj News", "BOL News", "Yaariyan Production", "Kids Land", "FACE of sialkot"}
seen, rows = set(), []
for r in csv.DictReader(open("search_raw.csv", encoding="utf-8")):
    if r["channel"] in CH and r["video_id"] not in seen: seen.add(r["video_id"]); rows.append(r)
out = []
with YoutubeDL({"quiet": True, "skip_download": True, "getcomments": False}) as y:
    for r in rows:
        try: d = y.extract_info(f"https://www.youtube.com/watch?v={r['video_id']}", download=False)
        except Exception as e: print("FAIL", r["video_id"], e); continue
        out.append(dict(channel=d.get("channel"), video_id=d["id"], upload_date=d.get("upload_date"), views=d.get("view_count"),
                        duration=d.get("duration"), title=d.get("title"), description=(d.get("description") or "")[:600].replace("\n", " ")))
        time.sleep(4)
with open("others.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(out[0])); w.writeheader(); w.writerows(out)
for o in sorted(out, key=lambda o: (o["channel"] or "", o["upload_date"] or "")):
    print(f"\n{o['channel']} | {o['upload_date']} | {o['views']} views | {o['duration']}s | {o['title']}\n   {o['description'][:300]}")
