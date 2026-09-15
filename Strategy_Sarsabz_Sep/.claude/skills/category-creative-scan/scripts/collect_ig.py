"""Collect Instagram profile stats + recent post creative via a headless browser.

No login, no cookie, no API key, no browser extension. Drives the browser-automation
skill's browser.mjs - a real headless Chromium, which Instagram serves the public
profile to.

    python collect_ig.py <workspace>/scan/handles.csv --out <workspace>/scan

Writes:
    profiles.csv        brand, handle, followers, following, verified, bio
    posts.csv           brand, handle, shortcode, date, caption, url, image_file, alt_text
    media/<handle>/     the post images themselves, <shortcode>.jpg

Pass --enrich to open each post for its TRUE post date and full caption. The profile
grid carries neither. Without real dates cadence is guesswork, so use it unless you are
only after images. Costs one extra page load per post.

Brands with a blank handle are carried through to profiles.csv as recorded absences -
a competitor with no account is a finding, not a hole in the sample.

KNOWN CEILING: Instagram serves ~12 posts to a logged-out browser and gates the rest
behind "Show more posts"; scrolling does not lift it (tested). Equal N per brand means
share of voice CANNOT be measured from this sample - never present one. For deeper
history or engagement numbers you need an authenticated session; see references/access.md.
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import re
import subprocess
import sys
import time
from pathlib import Path

import urllib.request

if hasattr(sys.stdout, "reconfigure"):        # Urdu alt text kills cp1252 on Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BROWSER = Path.home() / ".claude" / "skills" / "browser-automation" / "browser.mjs"

# Runs inside the page. Instagram serves 12 posts to a logged-out browser and then puts
# a "Show more posts" button under the grid - which DOES paginate without a login. Click
# it, scroll, repeat, until the target is met or the count stops growing.
PAGE_JS_TEMPLATE = r"""
(async () => {
  const TARGET = __TARGET__;
  const seenHrefs = () => new Set([...document.querySelectorAll('a[href*="/p/"], a[href*="/reel/"]')]
    .map(a => a.getAttribute("href")));
  const settle = async (n) => {
    for (let i = 0; i < n; i++) {
      window.scrollBy(0, 2500);
      await new Promise(r => setTimeout(r, 700));
    }
  };
  await settle(4);
  let last = 0;
  for (let round = 0; round < 30 && seenHrefs().size < TARGET; round++) {
    const btn = [...document.querySelectorAll('button, div[role="button"], a')]
      .find(b => /show more posts/i.test(b.textContent || ""));
    if (btn) { btn.click(); await new Promise(r => setTimeout(r, 1800)); }
    await settle(5);
    const now = seenHrefs().size;
    if (now === last && !btn) break;   // no button and no growth: genuinely the end
    if (now === last && round > 2) break;
    last = now;
  }
  const txt = document.body.innerText;
  const num = (re) => { const m = txt.match(re); return m ? m[1] : ""; };

  const seen = new Set();
  const posts = [];
  for (const a of document.querySelectorAll('a[href*="/p/"], a[href*="/reel/"]')) {
    const href = a.getAttribute("href") || "";
    const m = href.match(/\/(p|reel)\/([^\/]+)/);
    if (!m || seen.has(m[2])) continue;
    seen.add(m[2]);
    const img = a.querySelector("img");
    posts.push({
      shortcode: m[2],
      kind: m[1],
      href: "https://www.instagram.com" + href,
      img: img ? img.src : "",
      alt: img ? (img.alt || "") : "",
      // /<creator>/p/<code>/ carries a handle; the bare /p/<code>/ form does not.
      // Without this guard every ordinary post reads as a repost from "p" or "reel".
      reposted_from: (function () {
        const seg = href.split("/").filter(Boolean)[0] || "";
        return (seg === "p" || seg === "reel") ? "" : seg;
      })(),
    });
  }
  return JSON.stringify({
    title: document.title,
    followers: num(/([\d.,KM]+)\s+followers/i),
    following: num(/([\d.,KM]+)\s+following/i),
    verified: !!document.querySelector('svg[aria-label="Verified"]'),
    bio: txt.slice(0, 700).replace(/\s+/g, " "),
    posts,
  });
})()
"""

PROFILE_COLS = ["brand", "ring", "handle", "followers", "following", "verified",
                "posts_seen", "bio", "profile_url", "collected_at"]
POST_COLS = ["brand", "ring", "handle", "shortcode", "kind", "date", "date_source",
             "grid_position", "pinned_guess", "caption", "source_url", "image_file",
             "alt_text", "reposted_from"]

MONTHS = ("january february march april may june july august september october "
          "november december").split()


def date_from_alt(alt: str) -> str:
    """Instagram alt text usually reads 'Photo by X on August 18, 2026...'.

    A free date with no extra page load. Less reliable than --enrich (not every post
    has it) but far better than falling back to the collection date, which makes every
    brand look like it posted on the same day.
    """
    m = re.search(r"on\s+([A-Z][a-z]+)\s+(\d{1,2}),\s*(\d{4})", alt or "")
    if not m:
        return ""
    try:
        mon = MONTHS.index(m.group(1).lower()) + 1
    except ValueError:
        return ""
    return f"{int(m.group(3)):04d}-{mon:02d}-{int(m.group(2)):02d}"

# Runs on an individual post page. The profile grid does NOT carry post dates or full
# captions - only the post page does, and it gives both without a login. Worth the extra
# page load: without a real date, cadence is guesswork.
POST_JS = r"""
(() => {
  const t = document.querySelector("time");
  const txt = document.body.innerText;
  const m = txt.match(/\n\s*(\d+[smhdw])\s*\n([\s\S]{0,1200}?)(?:\n\s*\d+\s*(?:likes?|comments?)|\nView all|\nLog in|$)/);
  return JSON.stringify({
    date: t ? (t.getAttribute("datetime") || "") : "",
    caption: m ? m[2].replace(/\s+/g, " ").trim() : "",
  });
})()
"""


def render(url: str, js: str, timeout: int) -> dict | None:
    """Render a page with browser.mjs and return the parsed --eval payload."""
    try:
        proc = subprocess.run(
            ["node", str(BROWSER), url, "--eval", js, "--timeout", "45000"],
            capture_output=True, text=True, timeout=timeout,
            encoding="utf-8", errors="replace",  # profiles carry Urdu; cp1252 dies on it
        )
    except subprocess.TimeoutExpired:
        print("    TIMEOUT rendering the page")
        return None

    for line in proc.stdout.splitlines():
        if line.startswith("eval"):
            raw = line[len("eval"):].strip()
            try:
                return json.loads(json.loads(raw))  # eval prints a JSON-encoded string
            except json.JSONDecodeError as exc:
                print(f"    could not parse payload: {exc}")
                return None
    print(f"    no eval output (exit {proc.returncode}); stderr: {proc.stderr[:180]}")
    return None


def fetch_image(url: str, dest: Path) -> bool:
    if dest.exists():
        return True
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=45) as r:
            data = r.read()
        if len(data) < 1024:            # a 1x1 placeholder, not creative
            return False
        dest.write_bytes(data)
        return True
    except Exception as exc:
        print(f"    ! image failed: {type(exc).__name__}")
        return False


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("handles", type=Path)
    ap.add_argument("--out", type=Path, default=Path("workspace/audit"))
    ap.add_argument("--timeout", type=int, default=180, help="per-profile seconds")
    ap.add_argument("--pause", type=float, default=6.0, help="seconds between brands")
    ap.add_argument("--no-media", action="store_true")
    ap.add_argument("--depth", type=int, default=12,
                    help="target posts per brand. 12 is what Instagram serves before the "
                         "'Show more posts' button; higher values paginate through it. "
                         "Deeper reading gives real cadence and volume - and unequal N "
                         "across brands is what finally makes presence measurable.")
    ap.add_argument("--append", action="store_true",
                    help="add to existing profiles.csv/posts.csv instead of overwriting "
                         "them - use when re-collecting a single brand")
    ap.add_argument("--enrich", action="store_true",
                    help="open each post for its true date and full caption "
                         "(one extra page load per post; cadence is guesswork without it)")
    args = ap.parse_args()

    if not BROWSER.exists():
        sys.exit(f"browser.mjs not found at {BROWSER}")

    out = args.out
    media_root = out / "media"
    media_root.mkdir(parents=True, exist_ok=True)

    with args.handles.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))

    todo = [r for r in rows if (r.get("handle") or "").strip()]
    skipped = [r for r in rows if not (r.get("handle") or "").strip()]
    print(f"{len(todo)} brand(s) to collect; {len(skipped)} with no handle "
          f"(recorded as absent).\n")

    mode = "a" if args.append else "w"
    prof_path, post_path = out / "profiles.csv", out / "posts.csv"
    prof_exists = args.append and prof_path.exists() and prof_path.stat().st_size > 0
    post_exists = args.append and post_path.exists() and post_path.stat().st_size > 0
    prof_f = prof_path.open(mode, encoding="utf-8", newline="")
    prof_w = csv.DictWriter(prof_f, fieldnames=PROFILE_COLS)
    if not prof_exists:
        prof_w.writeheader()
    post_f = post_path.open(mode, encoding="utf-8", newline="")
    post_w = csv.DictWriter(post_f, fieldnames=POST_COLS)
    if not post_exists:
        post_w.writeheader()

    today = time.strftime("%Y-%m-%d")
    total_posts = total_imgs = 0

    for row in todo:
        brand, ring = row["brand"], row.get("ring", "")
        handle = row["handle"].strip().lstrip("@")
        print(f"--- {brand}  (@{handle})")

        page_js = PAGE_JS_TEMPLATE.replace("__TARGET__", str(args.depth))
        data = render(f"https://www.instagram.com/{handle}/", page_js, args.timeout)
        if not data:
            print()
            continue

        posts = data.get("posts", [])
        print(f"    {data.get('followers','?')} followers | {len(posts)} post(s) | "
              f"verified={data.get('verified')}")
        if not posts or not data.get("followers"):
            # A wrong-but-plausible handle returns the login wall, not an error. Left
            # unflagged it lands in profiles.csv as a normal row and ships as a finding.
            print(f"    ** WARNING: @{handle} returned no posts/followers. Either the "
                  f"handle is wrong or the account is private. VERIFY before using.")

        prof_w.writerow({
            "brand": brand, "ring": ring, "handle": handle,
            "followers": data.get("followers", ""), "following": data.get("following", ""),
            "verified": data.get("verified", ""), "posts_seen": len(posts),
            "bio": data.get("bio", "")[:400],
            "profile_url": f"https://www.instagram.com/{handle}/",
            "collected_at": today,
        })
        prof_f.flush()

        brand_dir = media_root / handle
        if not args.no_media and posts:
            brand_dir.mkdir(parents=True, exist_ok=True)

        for idx, p in enumerate(posts):
            image_file = ""
            if not args.no_media and p.get("img"):
                dest = brand_dir / f"{p['shortcode']}.jpg"
                if fetch_image(p["img"], dest):
                    image_file = str(dest.relative_to(out)).replace("\\", "/")
                    total_imgs += 1
            post_date, caption, date_source = "", "", ""
            if args.enrich:
                extra = render(p["href"], POST_JS, args.timeout) or {}
                post_date = (extra.get("date") or "")[:10]
                caption = (extra.get("caption") or "")[:900]
                if post_date:
                    date_source = "post_page"
                time.sleep(random.uniform(2, 4))  # jittered; do not remove
            if not post_date:
                post_date = date_from_alt(p.get("alt") or "")
                date_source = "alt_text" if post_date else ""
            # NEVER fall back to the collection date - it would make every brand look
            # like it posted on the same day and quietly destroy the cadence table.

            # a repost sits under the creator's handle, not the brand's
            reposted = p.get("reposted_from", "")
            post_w.writerow({
                "brand": brand, "ring": ring, "handle": handle,
                "shortcode": p["shortcode"], "kind": p.get("kind", ""),
                "date": post_date, "date_source": date_source,
                "grid_position": idx, "pinned_guess": "yes" if idx < 3 else "no",
                "caption": caption,
                "source_url": p.get("href", ""), "image_file": image_file,
                "alt_text": (p.get("alt") or "").replace("\n", " ")[:500],
                "reposted_from": "" if reposted == handle else reposted,
            })
            total_posts += 1
        post_f.flush()
        print()
        time.sleep(args.pause)

    for row in skipped:
        prof_w.writerow({
            "brand": row["brand"], "ring": row.get("ring", ""), "handle": "",
            "followers": "", "following": "", "verified": "", "posts_seen": 0,
            "bio": "NO INSTAGRAM ACCOUNT FOUND - " + (row.get("notes") or ""),
            "profile_url": "", "collected_at": today,
        })

    prof_f.close()
    post_f.close()
    print(f"Done. {total_posts} post(s), {total_imgs} image(s).")
    print(f"  {out/'profiles.csv'}\n  {out/'posts.csv'}\n  {media_root}")


if __name__ == "__main__":
    main()
