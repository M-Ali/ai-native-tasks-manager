"""Fetch public YouTube watch-page metadata (no API key): title, channel, publish date, views, length, description."""
import csv, json, re, sys, urllib.request
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
def vid(u):
    m = re.search(r"(?:youtu\.be/|v=)([\w-]{11})", u); return m.group(1) if m else None
def meta(v):
    html = urllib.request.urlopen(urllib.request.Request(f"https://www.youtube.com/watch?v={v}&hl=en", headers=UA), timeout=60).read().decode("utf-8", "replace")
    m = re.search(r"ytInitialPlayerResponse\s*=\s*(\{.+?\});(?:var|</script>)", html)
    if not m: return {"video_id": v, "status": "no_player_response"}
    p = json.loads(m.group(1)); d = p.get("videoDetails", {}); mf = p.get("microformat", {}).get("playerMicroformatRenderer", {})
    return {"video_id": v, "status": p.get("playabilityStatus", {}).get("status"), "title": d.get("title"), "channel": d.get("author"),
            "published": mf.get("publishDate", "")[:10], "views": d.get("viewCount"), "length_sec": d.get("lengthSeconds"),
            "description": (d.get("shortDescription") or "").replace("\n", " ")[:600]}
if __name__ == "__main__":
    rows = list(csv.DictReader(open(sys.argv[1], encoding="utf-8")))
    out = []
    for r in rows:
        v = vid(r["url"])
        try: r.update(meta(v))
        except Exception as e: r.update({"video_id": v, "status": f"error {e}"})
        out.append(r); print(r.get("published"), r.get("views"), r.get("length_sec"), r.get("channel"), "|", r.get("title"), file=sys.stderr)
    w = csv.DictWriter(open(sys.argv[2], "w", newline="", encoding="utf-8"), fieldnames=list(out[0].keys())); w.writeheader(); w.writerows(out)
