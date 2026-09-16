"""Salam Kissan YouTube comments -> comment_codes.csv + counts printed for comment_findings.md.

Base: corpus.csv, built from raw/*.info.json (yt-dlp), 18 Salam Kissan uploads with >=380,000 views.
Every comment was read (partial_read.txt, partial_read2.txt, and the retried films).

Two kinds of code, kept apart on purpose:
- PATTERN codes: a regex; the count is whatever it matches, and the matches are listed for checking.
- VERIFIED codes: a list of verbatim fragments picked by reading. Each must match exactly one comment
  or the script stops. These carry the judgement ("is this a complaint?") that a regex cannot.
"""
import csv, re, sys
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).parent
rows = list(csv.DictReader((HERE / "corpus.csv").open(encoding="utf-8")))
aud = [r for r in rows if r["is_uploader"] == "no"]

PATTERNS = {
    "channel_winner_notice": (rows, lambda r: r["is_uploader"] == "yes" and re.search(r"(?i)email your complete details", r["text"])),
    "vlog_referral": (aud, lambda r: re.search(r"(?i)sh[ae]h?e?r?r? ?m[ae]?[iy]?n? ?d[ei]h?[iy]?a?a?h?t|sher ?m[ae]?[iy]?n? ?d[ei]|shahr ma dihat|ducky|turab", r["text"])),
    "subscriber_100k": (aud, lambda r: re.search(r"(?i)100 ?k", r["text"])),
    "contest_own_words": (aud, lambda r: re.search(r"(?i)wish to win|hope to win|gift i received|gift if i won|details sent|^\W*done\W*$|gify milna", r["text"])),
    "india_bangladesh": (aud, lambda r: re.search(r"(?i)\bindia|bangladesh|jai (jawan|javan|johar)|modi\b", r["text"])),
    "song_voice_lyrics": (aud, lambda r: re.search(r"(?i)\bsong|lyric|\bvoice|singer|\bmusic|noori|ali noor|mai dha|گانا|آواز", r["text"])),
    "salam_pakistan": (aud, lambda r: re.search(r"(?i)salam,? ?(ka )?pakistan|سلام پاکستان", r["text"])),
    "fraud_gang_replies": (aud, lambda r: r["text"].strip() == "Fatima Group is Fraud Gang"),
}

# Read, then picked. Each fragment must hit exactly one comment.
VERIFIED = {
    "grievance": [  # a complaint about farmers' conditions, prices, supply or neglect - not praise that mentions hardship
        "کسانوں کو کھاد نہیں مل رہی",
        "کھاد کے لیے پریشان ہیں",
        "Government should assist them financially",
        "کبھی غریب کسان کو حق نہیں دیتے",
        "Farmers must be appreciated and helped financially",
        "Provide them facilities",
        "gareebon ka heq kha gae",
        "کاش اس ملک میں کسان کی قدر ھوتی",
        "جہاں کسان کی کوئی عزت نا ہونے کے برابر ہے",
        "حکومت مونگ تہ دہ تخمونہ ارزان",
        "Khuda kissano ko on ka haq milay",
        "kisi hakomat nay humayn izat nhi di",
        "khaad nayaab ho cuki hy",
        "کاش اس ملک میں کسان کی قدر ہوتی",
        "DAP price 9000 tk",
        "should not be black marketed",
        "sasti khaad, acha beej, fixed market prices",
        "koi punjabi farmers ka ehsaas nai krta",
        "کاش کسان کی زندگی بھی کوئی حکومت کرتی آسان",
    ],
}

def one(fragment):
    hits = [r for r in rows if fragment in r["text"]]
    if len(hits) != 1:
        sys.exit(f"fragment matched {len(hits)} comments, need exactly 1: {fragment!r}")
    return hits[0]

codes = defaultdict(set)
for name, (pool, test) in PATTERNS.items():
    for r in pool:
        if test(r): codes[r["comment_id"]].add(name)
for name, frags in VERIFIED.items():
    for f in frags: codes[one(f)["comment_id"]].add(name)

with (HERE / "comment_codes.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh); w.writerow(["comment_id", "video_id", "upload_date", "author", "is_uploader", "is_reply", "likes", "codes", "text"])
    for r in rows:
        w.writerow([r["comment_id"], r["video_id"], r["upload_date"], r["author"], r["is_uploader"], r["is_reply"], r["likes"],
                    ";".join(sorted(codes.get(r["comment_id"], ()))), r["text"]])

print(f"BASE: {len(rows)} comments | channel {len(rows) - len(aud)} | audience {len(aud)} | "
      f"top-level {sum(r['is_reply'] == 'no' for r in aud)} | authors {len({r['author'] for r in aud})}")
print("\nPER FILM (all comments / audience / channel)")
films = defaultdict(list)
for r in rows: films[(r["upload_date"], r["video_id"], r["video_title"])].append(r)
for (d, v, t), rs in sorted(films.items()):
    print(f"  {d} {v} {len(rs):>4} {sum(x['is_uploader']=='no' for x in rs):>4} {sum(x['is_uploader']=='yes' for x in rs):>3}  {t[:55]}")

print("\nCODE COUNTS (comments; films)")
for name in list(PATTERNS) + list(VERIFIED):
    hit = [r for r in rows if name in codes.get(r["comment_id"], ())]
    print(f"  {name:24} {len(hit):>4}  films={dict(Counter(r['video_id'] for r in hit))}")

print("\n2025 SHORTS: comments by authors who posted on 5+ of the eight 2025 uploads")
y25 = [r for r in aud if r["upload_date"].startswith("2025")]
fa = defaultdict(set)
for r in y25: fa[r["author"]].add(r["video_id"])
multi = {a for a, vs in fa.items() if len(vs) >= 5}
print(f"  2025 audience comments {len(y25)} | authors {len(fa)} | on 5+ films: {sorted(multi)} -> {sum(r['author'] in multi for r in y25)} comments")

for name in ("vlog_referral", "contest_own_words", "subscriber_100k", "india_bangladesh"):
    print(f"\n== {name}")
    for r in rows:
        if name in codes.get(r["comment_id"], ()): print(f"  [{r['video_id'][:5]}] ({r['likes']}) {r['text'][:120]!r}")
