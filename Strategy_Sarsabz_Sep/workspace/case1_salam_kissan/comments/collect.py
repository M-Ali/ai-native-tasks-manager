"""Collect YouTube comments on Sarsabz's biggest Salam Kissan / Kissan Day films -> raw/*.info.json + comments.csv.

Selection rule (stated because it shapes the corpus): every Salam Kissan / Kissan Day upload on @Sarsabz with
>= 380,000 views - the films the brand put its media weight behind. Paced to avoid YouTube's bot check.
"""
import csv, json, re, subprocess, sys, time
from pathlib import Path
HERE = Path(__file__).parent
INV = HERE.parents[1] / "case2_audit" / "owned" / "channel_inventory.csv"
pat = re.compile(r"(?i)salam ?kiss?an|kiss?an day|behind every meal|kissan ki mehnat|jafakash")
rows = [r for r in csv.DictReader(INV.open(encoding="utf-8"))
        if pat.search((r["title"] or "") + " " + (r.get("description") or ""))
        and int(r["views_exact"] or r["views"] or 0) >= 380_000]
rows.sort(key=lambda r: r["published"])
(HERE / "raw").mkdir(exist_ok=True)
with (HERE / "films.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh); w.writerow(["video_id", "published", "views", "title"])
    for r in rows: w.writerow([r["video_id"], r["published"], r["views_exact"] or r["views"], r["title"]])
print(len(rows), "films selected", flush=True)
for r in rows:
    vid = r["video_id"]
    if (HERE / "raw" / f"{vid}.info.json").exists():
        continue
    url = f"https://www.youtube.com/watch?v={vid}" if r["tab"] == "videos" else f"https://www.youtube.com/shorts/{vid}"
    res = subprocess.run([sys.executable, "-m", "yt_dlp", "--skip-download", "--write-comments", "--no-write-thumbnail",
                          "-o", str(HERE / "raw" / "%(id)s"), "--extractor-args", "youtube:max_comments=3000,all,200",
                          "--sleep-requests", "1", "-q", "--no-warnings", url], capture_output=True, text=True)
    status = "ok" if (HERE / "raw" / f"{vid}.info.json").exists() else ("BLOCKED" if "bot" in res.stderr else "failed")
    print(f"{r['published']} {vid} {status} {res.stderr.strip()[:120]}", flush=True)
    if status == "BLOCKED":
        break
    time.sleep(8)
out = []
for f in sorted((HERE / "raw").glob("*.info.json")):
    d = json.load(f.open(encoding="utf-8"))
    for c in d.get("comments") or []:
        out.append({"video_id": d["id"], "video_title": d.get("title", ""), "text": c.get("text", ""),
                    "likes": c.get("like_count") or 0, "is_reply": "yes" if c.get("parent") not in (None, "root") else "no",
                    "timestamp": c.get("timestamp") or ""})
with (HERE / "comments.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=["video_id", "video_title", "text", "likes", "is_reply", "timestamp"])
    w.writeheader(); w.writerows(out)
print(f"DONE: {len(out)} comments from {len(list((HERE / 'raw').glob('*.info.json')))} films -> comments.csv", flush=True)
