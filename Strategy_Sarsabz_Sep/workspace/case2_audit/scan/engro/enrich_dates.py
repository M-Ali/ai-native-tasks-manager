"""Replace approximate dates with true publish dates (+ exact views, length, description) from each watch page."""
import csv, sys
from concurrent.futures import ThreadPoolExecutor
sys.path.insert(0, ".")
from yt_meta import meta  # noqa
rows = list(csv.DictReader(open("channel_inventory.csv", encoding="utf-8")))
def work(r):
    for _ in range(3):
        try:
            m = meta(r["video_id"])
            if m.get("published"):
                return {**r, "published": m["published"], "views_exact": m.get("views"), "length_exact": m.get("length_sec"),
                        "status": m.get("status"), "description": m.get("description")}
        except Exception as e:
            err = str(e)
    return {**r, "published": "", "views_exact": "", "length_exact": "", "status": "failed", "description": ""}
with ThreadPoolExecutor(8) as ex:
    out = list(ex.map(work, rows))
with open("channel_inventory.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(out[0])); w.writeheader(); w.writerows(out)
print("failed:", sum(1 for r in out if not r["published"]), "of", len(out))
