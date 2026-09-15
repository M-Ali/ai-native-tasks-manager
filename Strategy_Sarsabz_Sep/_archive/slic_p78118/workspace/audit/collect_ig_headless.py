"""Collect Instagram profile stats + recent post creative via a headless browser.

No login, no cookie, no API key. Drives the browser-automation skill's browser.mjs,
which renders Instagram as a real Chromium and so gets served the public profile.

    uv run python workspace/audit/collect_ig_headless.py workspace/audit/handles.csv \
      --out workspace/audit

Writes:
    profiles.csv        brand, handle, followers, following, verified, bio
    posts.csv           brand, handle, shortcode, url, image_file, alt_text
    media/<handle>/     the post images themselves, <shortcode>.jpg

KNOWN CEILING: Instagram serves ~12 posts to a logged-out browser and gates the
rest behind "Show more posts". Scrolling does not lift it (tested). For deeper
history you need an authenticated session - see collect_instagram.py.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
import time
from pathlib import Path

import urllib.request

BROWSER = Path.home() / ".claude" / "skills" / "browser-automation" / "browser.mjs"

# Runs inside the page. Scrolls to trigger lazy-load, then reads the grid.
PAGE_JS = r"""
(async () => {
  for (let i = 0; i < 5; i++) {
    window.scrollBy(0, 1500);
    await new Promise(r => setTimeout(r, 800));
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
      reposted_from: href.split("/")[1] || "",
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
POST_COLS = ["brand", "ring", "handle", "shortcode", "kind", "source_url",
             "image_file", "alt_text", "reposted_from"]


def run_page(handle: str, timeout: int) -> dict | None:
    """Render one profile and return the parsed payload, or None on failure."""
    url = f"https://www.instagram.com/{handle}/"
    try:
        proc = subprocess.run(
            ["node", str(BROWSER), url, "--eval", PAGE_JS, "--timeout", "45000"],
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

    prof_f = (out / "profiles.csv").open("w", encoding="utf-8", newline="")
    prof_w = csv.DictWriter(prof_f, fieldnames=PROFILE_COLS)
    prof_w.writeheader()
    post_f = (out / "posts.csv").open("w", encoding="utf-8", newline="")
    post_w = csv.DictWriter(post_f, fieldnames=POST_COLS)
    post_w.writeheader()

    today = time.strftime("%Y-%m-%d")
    total_posts = total_imgs = 0

    for row in todo:
        brand, ring = row["brand"], row.get("ring", "")
        handle = row["handle"].strip().lstrip("@")
        print(f"--- {brand}  (@{handle})")

        data = run_page(handle, args.timeout)
        if not data:
            print()
            continue

        posts = data.get("posts", [])
        print(f"    {data.get('followers','?')} followers | {len(posts)} post(s) | "
              f"verified={data.get('verified')}")

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

        for p in posts:
            image_file = ""
            if not args.no_media and p.get("img"):
                dest = brand_dir / f"{p['shortcode']}.jpg"
                if fetch_image(p["img"], dest):
                    image_file = str(dest.relative_to(out)).replace("\\", "/")
                    total_imgs += 1
            # a repost sits under the creator's handle, not the brand's
            reposted = p.get("reposted_from", "")
            post_w.writerow({
                "brand": brand, "ring": ring, "handle": handle,
                "shortcode": p["shortcode"], "kind": p.get("kind", ""),
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
