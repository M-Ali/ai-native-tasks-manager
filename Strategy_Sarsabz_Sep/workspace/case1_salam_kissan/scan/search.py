"""YouTube search: who publishes about Kissan Day / National Farmers Day Pakistan. Flat metadata only, paced.
Writes search_raw.csv (query, rank, video_id, channel, title, views, duration). Search rank is YouTube's, not a market measure."""
import csv, time, sys
from yt_dlp import YoutubeDL
sys.stdout.reconfigure(encoding="utf-8")
Q = ["Kissan Day 18 December", "National Kissan Day Pakistan", "Farmers Day Pakistan 18 December",
     "کسان ڈے", "Salam Kissan", "Kisan Day Pakistan", "Happy Kissan Day"]
rows = []
with YoutubeDL({"quiet": True, "extract_flat": True, "skip_download": True}) as y:
    for q in Q:
        try:
            info = y.extract_info(f"ytsearch60:{q}", download=False)
        except Exception as e:
            print("FAIL", q, e); continue
        for i, e in enumerate(info.get("entries") or [], 1):
            rows.append(dict(query=q, rank=i, video_id=e.get("id"), channel=e.get("channel") or e.get("uploader"),
                             title=e.get("title"), views=e.get("view_count"), duration=e.get("duration")))
        print(q, len(info.get("entries") or [])); time.sleep(8)
with open("search_raw.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
print("rows", len(rows))
