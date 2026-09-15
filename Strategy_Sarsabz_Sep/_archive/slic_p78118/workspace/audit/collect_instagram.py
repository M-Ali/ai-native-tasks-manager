"""Collect Instagram creative + metrics for the competitor comms audit.

Needs an authenticated instaloader session. Anonymous access is refused with
HTTP 429 (tested 2026-08-27), so there is no session-free path.

Setup, once (run this yourself - it reads your browser cookie jar):

    uv run --with instaloader --with browser_cookie3 \
      instaloader --load-cookies chrome --sessionfile workspace/audit/ig.session :stories

Then:

    uv run --with instaloader python workspace/audit/collect_instagram.py \
      workspace/audit/handles.csv --sessionfile workspace/audit/ig.session \
      --user YOUR_IG_USERNAME --out workspace/audit --since 2025-09-01

Writes:
    profiles.csv          brand, handle, followers, mediacount, verified, bio
    posts.csv             one row per post, capture.xlsx column names where they match
    media/<handle>/       every image and video file, named <shortcode>_<n>.<ext>

Re-runnable. Posts already downloaded are skipped, so a rate-limit stop can be
resumed by running the same command again.
"""

from __future__ import annotations

import argparse
import csv
import random
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import instaloader
from instaloader.exceptions import (
    ConnectionException,
    LoginRequiredException,
    ProfileNotExistsException,
    TooManyRequestsException,
)

PROFILE_COLS = [
    "brand", "ring", "handle", "followers", "followees", "mediacount",
    "is_verified", "is_business", "category", "biography", "external_url",
    "collected_at",
]

POST_COLS = [
    "brand", "ring", "handle", "shortcode", "date_utc", "format", "is_video",
    "duration_sec", "likes", "comments", "video_view_count", "caption",
    "hashtags", "mentions", "is_sponsored", "media_files", "source_url",
]


def pause(lo: float, hi: float) -> None:
    """Jittered sleep. Instagram rate-limits hard; do not remove this."""
    time.sleep(random.uniform(lo, hi))


def post_format(post: instaloader.Post) -> str:
    """Map to the capture.xlsx `format` vocabulary."""
    if post.typename == "GraphSidecar":
        return "carousel"
    if post.is_video:
        return "video"
    return "static"


def read_handles(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = [r for r in csv.DictReader(fh) if (r.get("handle") or "").strip()]
    if not rows:
        sys.exit(f"{path}: no rows with a handle. Fill the handle column first.")
    return rows


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("handles", type=Path, help="CSV with brand,ring,handle")
    ap.add_argument("--sessionfile", type=Path, required=True)
    ap.add_argument("--user", required=True, help="the Instagram username the session belongs to")
    ap.add_argument("--out", type=Path, default=Path("workspace/audit"))
    ap.add_argument("--since", default="2025-09-01",
                    help="stop walking a profile once posts predate this (YYYY-MM-DD)")
    ap.add_argument("--limit", type=int, default=40, help="max posts per brand")
    ap.add_argument("--no-media", action="store_true", help="metadata only, skip image/video files")
    args = ap.parse_args()

    since = datetime.strptime(args.since, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    out = args.out
    media_root = out / "media"
    media_root.mkdir(parents=True, exist_ok=True)

    L = instaloader.Instaloader(
        quiet=True,
        download_comments=False,
        save_metadata=False,
        compress_json=False,
        dirname_pattern=str(media_root / "{target}"),
        filename_pattern="{shortcode}",
        post_metadata_txt_pattern="",
    )
    try:
        L.load_session_from_file(args.user, str(args.sessionfile))
    except FileNotFoundError:
        sys.exit(f"No session at {args.sessionfile}. Run the --load-cookies step in the docstring first.")
    if not L.test_login():
        sys.exit("Session did not authenticate. Re-run the --load-cookies step.")

    brands = read_handles(args.handles)
    print(f"Session OK as @{args.user}. {len(brands)} brand(s) to collect.\n")

    profiles_path, posts_path = out / "profiles.csv", out / "posts.csv"
    seen: set[str] = set()
    if posts_path.exists():
        with posts_path.open(encoding="utf-8", newline="") as fh:
            seen = {r["shortcode"] for r in csv.DictReader(fh) if r.get("shortcode")}
        print(f"Resuming - {len(seen)} post(s) already collected.\n")

    pf = profiles_path.open("a", encoding="utf-8", newline="")
    pw = csv.DictWriter(pf, fieldnames=PROFILE_COLS)
    if profiles_path.stat().st_size == 0:
        pw.writeheader()
    sf = posts_path.open("a", encoding="utf-8", newline="")
    sw = csv.DictWriter(sf, fieldnames=POST_COLS)
    if posts_path.stat().st_size == 0:
        sw.writeheader()

    for row in brands:
        brand, ring, handle = row["brand"], row.get("ring", ""), row["handle"].strip().lstrip("@")
        print(f"--- {brand}  (@{handle})")
        try:
            p = instaloader.Profile.from_username(L.context, handle)
        except ProfileNotExistsException:
            print("    handle does not exist - check it and re-run\n")
            continue
        except (TooManyRequestsException, ConnectionException) as exc:
            print(f"    RATE LIMITED: {exc}\n    Stopping. Wait a few hours, then re-run to resume.")
            break
        except LoginRequiredException:
            sys.exit("    Session expired mid-run. Re-import cookies and re-run.")

        pw.writerow({
            "brand": brand, "ring": ring, "handle": handle,
            "followers": p.followers, "followees": p.followees,
            "mediacount": p.mediacount, "is_verified": p.is_verified,
            "is_business": p.is_business_account, "category": p.business_category_name or "",
            "biography": (p.biography or "").replace("\n", " "),
            "external_url": p.external_url or "",
            "collected_at": datetime.now(timezone.utc).date().isoformat(),
        })
        pf.flush()
        print(f"    {p.followers:,} followers | {p.mediacount:,} posts | verified={p.is_verified}")

        kept = 0
        try:
            for post in p.get_posts():
                if kept >= args.limit:
                    break
                if post.date_utc.replace(tzinfo=timezone.utc) < since:
                    print(f"    reached {args.since}, stopping this brand")
                    break
                if post.shortcode in seen:
                    continue

                files: list[str] = []
                if not args.no_media:
                    try:
                        L.download_post(post, target=handle)
                        files = sorted(
                            str(f.relative_to(out))
                            for f in (media_root / handle).glob(f"{post.shortcode}*")
                            if f.suffix.lower() in {".jpg", ".jpeg", ".png", ".mp4", ".webp"}
                        )
                    except Exception as exc:  # a single bad asset must not kill the run
                        print(f"    ! media failed for {post.shortcode}: {type(exc).__name__}")

                sw.writerow({
                    "brand": brand, "ring": ring, "handle": handle,
                    "shortcode": post.shortcode,
                    "date_utc": post.date_utc.date().isoformat(),
                    "format": post_format(post), "is_video": post.is_video,
                    "duration_sec": getattr(post, "video_duration", "") or "",
                    "likes": post.likes, "comments": post.comments,
                    "video_view_count": post.video_view_count or "",
                    "caption": (post.caption or "").replace("\n", " "),
                    "hashtags": " ".join(post.caption_hashtags),
                    "mentions": " ".join(post.caption_mentions),
                    "is_sponsored": post.is_sponsored,
                    "media_files": " | ".join(files),
                    "source_url": f"https://www.instagram.com/p/{post.shortcode}/",
                })
                sf.flush()
                seen.add(post.shortcode)
                kept += 1
                pause(3, 7)
        except (TooManyRequestsException, ConnectionException) as exc:
            print(f"    RATE LIMITED after {kept} post(s): {exc}")
            print("    Stopping. Wait a few hours, then re-run to resume.")
            break

        print(f"    {kept} post(s) collected\n")
        pause(20, 40)

    pf.close()
    sf.close()
    print(f"Done. profiles.csv and posts.csv in {out}; media in {media_root}")


if __name__ == "__main__":
    main()
