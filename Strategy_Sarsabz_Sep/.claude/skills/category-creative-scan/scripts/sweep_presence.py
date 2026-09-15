"""Follower sweep: where does each brand's audience actually live?

Probes a candidate URL per brand per platform in a logged-out headless browser and
records the follower/subscriber count, the page name and a bio snippet.

    python sweep_presence.py candidates.csv --out presence.csv

Why this exists: the Instagram scan found PTCL at 91K followers. Its Facebook page has
1.3M. Fiberlink was recorded as "no account" because it has no Instagram - it has 23K on
Facebook. Choosing a scan platform without checking audience size first understates the
category, and can call a present brand absent.

Two calls per URL, not one. Facebook and TikTok both serve an interstitial or a WAF
challenge on first load; the second call against the same --session lands on the real
page. See references/access.md.

A tiny follower count for a national operator is a FAILED LOOKUP, not a finding. The
status column says `suspect` for those so they get searched again rather than published.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BROWSER = Path.home() / ".claude" / "skills" / "browser-automation" / "browser.mjs"
SESSION = "sweep"

PAGE_JS = r"""
(async () => {
  const close = document.querySelector('[aria-label=Close],[aria-label="close"]');
  if (close) { try { close.click(); } catch (e) {} }
  await new Promise(r => setTimeout(r, 2500));
  const t = document.body.innerText || "";
  return { title: document.title.slice(0, 90), text: t.slice(0, 1200).replace(/\s+/g, " ") };
})()
"""

# order matters: "subscribers" before the generic follower patterns
COUNT_PATTERNS = [
    re.compile(r"([\d.,]+\s*[KM]?)\s*subscribers", re.I),
    re.compile(r"([\d.,]+\s*[KM]?)\s*Followers", re.I),
    re.compile(r"([\d.,]+\s*[KM]?)\s*followers", re.I),
]


def to_number(raw: str) -> int | None:
    """'1.3M' -> 1300000. Returns None rather than guessing on an odd format."""
    s = raw.strip().replace(",", "").upper()
    m = re.fullmatch(r"([\d.]+)\s*([KM]?)", s)
    if not m:
        return None
    try:
        val = float(m.group(1))
    except ValueError:
        return None
    return int(val * {"": 1, "K": 1_000, "M": 1_000_000}[m.group(2)])


def run(args: list[str], timeout: int = 90) -> str:
    try:
        p = subprocess.run([*args], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        return (p.stdout or "") + (p.stderr or "")
    except subprocess.TimeoutExpired:
        return "TIMEOUT"


def probe(url: str) -> dict:
    run(["node", str(BROWSER), url, "--session", SESSION, "--timeout", "45000"])
    out = run(["node", str(BROWSER), "--session", SESSION, "--eval", PAGE_JS,
               "--timeout", "45000"])
    m = re.search(r"^eval\s+(.*)$", out, re.M)
    if not m:
        return {"status": "no_eval", "title": "", "text": ""}
    try:
        data = json.loads(m.group(1).strip())
    except json.JSONDecodeError:
        return {"status": "bad_json", "title": "", "text": m.group(1)[:200]}
    if isinstance(data, str):                       # "EVAL ERROR: ..."
        return {"status": "eval_error", "title": "", "text": data[:200]}
    return {"status": "ok", **data}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("candidates", type=Path)
    ap.add_argument("--out", type=Path, default=Path("presence.csv"))
    ap.add_argument("--suspect-below", type=int, default=500,
                    help="follower count below which a national brand's lookup is "
                         "flagged `suspect` rather than published as fact")
    args = ap.parse_args()

    if not BROWSER.exists():
        sys.exit(f"browser.mjs not found at {BROWSER}")

    with args.candidates.open(encoding="utf-8-sig", newline="") as fh:
        cands = [r for r in csv.DictReader(fh) if (r.get("url") or "").strip()]

    rows = []
    for i, c in enumerate(cands, 1):
        brand, platform, url = c["brand"], c["platform"], c["url"].strip()
        print(f"[{i:>2}/{len(cands)}] {brand} - {platform}")
        res = probe(url)
        text = res.get("text", "")
        raw = count = None
        for pat in COUNT_PATTERNS:
            m = pat.search(text)
            if m:
                raw = m.group(1).strip()
                count = to_number(raw)
                break

        status = res["status"]
        if status == "ok":
            if count is None:
                status = "no_count"
            elif count < args.suspect_below:
                status = "suspect"
        print(f"          {status:<10} {raw or '-':>8}  {res.get('title','')[:56]}")

        rows.append({"brand": brand, "platform": platform, "url": url,
                     "followers_raw": raw or "", "followers": count if count is not None else "",
                     "status": status, "title": res.get("title", ""),
                     "snippet": text[:220]})

    run(["node", str(BROWSER), "--session", SESSION, "--close"], timeout=30)

    with args.out.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["brand", "platform", "url", "followers_raw",
                                           "followers", "status", "title", "snippet"])
        w.writeheader()
        w.writerows(rows)
    ok = sum(1 for r in rows if r["status"] == "ok")
    print(f"\n{args.out}: {len(rows)} probe(s), {ok} with a verified count")
    print("`suspect` and `no_count` rows need a handle search before they are published.")


if __name__ == "__main__":
    main()
