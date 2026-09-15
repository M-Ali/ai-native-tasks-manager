# Getting at the creative

What works, what does not, and what each route costs. Tested 27 August 2026 — platform
access changes, so re-test rather than trusting this list, and update it when it moves.

## Instagram

### What works: headless browser (the default)

`scripts/collect_ig.py` drives `browser.mjs` from the `browser-automation` skill — a real
headless Chromium. Instagram serves it the public profile.

Yields per profile: follower count, following, bio, verified flag, and ~12 recent posts
with shortcodes, post URLs, alt text, and **downloadable CDN image URLs at 640px**.

No login, no cookie, no API key, no payment, no browser extension.

**Post dates and full captions need one extra hop.** The profile grid carries neither.
Opening each post URL does — `<time datetime>` gives an exact timestamp and the body
carries the full caption, both without a login. `collect_ig.py --enrich` does this.
Use it unless you only want images: without real dates, cadence is guesswork.

**Likes and view counts are NOT available logged out.** Tested on three posts (one reel,
two static) — all null. If a source claims otherwise, verify before relying on it;
engagement needs an authenticated session.

**The ceiling is 24 posts, not 12.** The grid loads 12, then shows a "Show more posts"
button which **does** paginate without a login — one click yields a second tranche of 12.
After that the button disappears and no amount of scrolling or clicking adds more (traced
round by round). So `--depth 24` is free; anything beyond 24 needs an authenticated
session.

Scrolling alone never lifts the first 12 — the button click is what does it. Consequences
of the 24 ceiling:

- Equal N per brand, so **share of voice is unmeasurable**. Never present one.
- Cadence is a rate over the recent window. State the window.
- Screenshots of the profile grid come out blank because tiles lazy-load. Do not
  screenshot the grid — read image URLs from the DOM and download them, which is what the
  collector does.

### What does not work

| Route | Result |
|---|---|
| `curl` / `requests` on the profile | Login-walled shell, ~615 KB, identical bytes for every account. No bio, no counts |
| `/p/<code>/embed/captioned/` | Returns HTTP 200 but the payload is login-walled. This used to work |
| `instaloader`, anonymous | **HTTP 429 on the first call.** Not throttling — anonymous access is refused. An 8-minute retry run ended the same way |

Do not spend time rediscovering these.

### Deeper history: authenticated session

Needed only when 12 posts per brand is genuinely insufficient, or when engagement numbers
(likes, comments, view counts) are required to weight the analysis.

```bash
# the user runs this themselves - it reads their browser cookie jar
uv run --with instaloader --with browser_cookie3 \
  instaloader --load-cookies chrome --sessionfile ig.session :stories
```

Then `instaloader --sessionfile ig.session profile <handle>` yields unlimited posts,
total post count, likes, comments and view counts.

Three cautions, and they matter:

- **Use a secondary account.** Automated collection breaches Instagram's ToS and accounts
  get action-blocked. The risk sits on whichever account holds the cookie.
- **Pace it.** Sequential brands, jittered sleeps. Expect hours for a large set.
- **Never read the user's cookie jar yourself.** Ask them to run the command. Reading a
  credential store on someone's behalf is not yours to do.

### Paid alternative

Apify's Instagram Scraper (`apify/instagram-scraper`) returns the same fields with no
risk to the user's account and no rate-limit babysitting. Roughly a couple of dollars for
a 16-brand job; a free monthly credit often covers it. Worth recommending when a deadline
is tight — check current pricing rather than quoting from memory.

## Facebook

**Partially open — more than previously recorded.** Re-tested 11 September 2026 with the
headless browser and the `--session` two-call pattern (see TikTok below). An
unauthenticated *server-side* fetch still returns the title and nothing else, but the
browser renders the page header reliably:

- follower and following counts, page name, verification
- the Intro/bio text, page category, address
- ~19-21 CDN images
- **one or two recent posts** — caption fragment, reaction count, even a comment

Then it stops. A `[role=dialog]` login wall persists through Escape and close-button
clicks, scrolling adds nothing, and the body text caps around 1,000 characters. So
Facebook is good for **presence, audience size and existence checks**, and unusable for a
systematic per-brand read of 12-24 posts. Do not plan a scan around it.

The workaround for specific copy is unchanged: a targeted web search on a distinctive
phrase recovers exact captions with a citable post URL.

**Check follower counts here before concluding where a category lives.** In the Pakistani
fixed-broadband scan, PTCL had 91K Instagram followers and **1.3M on Facebook**; Fiberlink
had **no Instagram at all and 23K on Facebook**. An Instagram-only scan would have called
one of those brands absent and understated the category's real centre of gravity by more
than tenfold. The follower sweep is cheap - do it before choosing the platform.

**But post text is indexed by search.** A targeted web search on a distinctive phrase
recovers exact captions with a citable post URL:

```
"BrandName" facebook "the distinctive phrase from the ad"
```

This recovered a verbatim competitor caption after the page itself refused to load. It
gives text, not completeness, and never images — but it is not nothing, and the citation
is real.

## Meta Ad Library

`facebook.com/ads/library` — the only public source of **paid** creative with dates an
evaluator can verify. Server-side fetching fails (JavaScript app; the connection drops).

It needs a logged-in browser, so it is usually a manual step for the user: set country,
set "All ads", search each brand, screenshot into `<workspace>/scan/adlibrary/`.

Two limits that are platform behaviour, not tooling gaps:

- **No spend is shown for commercial advertisers.** Only political and issue ads carry
  spend. Never present an estimated spend figure as measured.
- **No export exists** for commercial advertisers.

Worth the manual effort on a pitch, because "here is what they are actually running in
paid, with dates" is the most verifiable evidence in a competitive section.

## Owned channels

Company campaign libraries are underrated: fully citable, often years deep, and they
carry taglines and asset links. One competitor's own site yielded a 19-year campaign
history with lines and video IDs.

The catch: it is the brand's curated view of itself, and coding from a page's *description*
of a campaign is not the same as coding from the film. Mark such rows `provisional` and
verify against the asset before anything leans on them.

## TikTok

Tested 11 September 2026. **Works logged out, but only with a persistent session.**

A single `browser.mjs <url> --eval` call fails: TikTok serves a WAF interstitial
("Please wait…"), and when it clears, the navigation destroys the eval context —
`Execution context was destroyed` is the challenge completing, not a bug in your JS.

The two-call pattern works:

```bash
node browser.mjs "https://www.tiktok.com/@<handle>" --session tt1   # takes the challenge
node browser.mjs --session tt1 --eval '<js>'                        # lands on the profile
node browser.mjs --session tt1 --close
```

Yields logged out: follower / following / likes counts, bio, video links, and —
unlike Instagram — **per-video view counts** on the grid tiles. That makes attention
measurable rather than merely presence, which is worth more than extra depth.
Scrolling loads further tiles; re-test the ceiling before promising a number.

## YouTube

A server-side fetch of `/videos` returns only the footer. **The headless browser renders
the channel page fine logged out** (tested 11 September 2026) — title, handle, subscriber
area and video links. Prefer the browser over `requests` here, as with Instagram.

Individual video pages fetch acceptably. Where a brand's campaign library links video
IDs, go directly to those rather than enumerating the channel.

### Both platforms: the handle problem is worse here than on Instagram

Guessing handles on TikTok fails loudly and silently at once. In one test every guessed
handle resolved to a real page that was **not the brand**: `@ptclofficial` (129 followers,
bio "LIVE POD CAST", zero videos), `@stormfiber` (3 followers), `@nayatel` (6 followers),
`@jazz` (a private personal account belonging to someone called Jazmine). On YouTube,
`@ptclofficial` is an empty channel — "This channel doesn't have any content" — while the
brand's real presence sits elsewhere.

A squatted or dormant namesake returns HTTP 200 and a plausible-looking profile. Search
for the handle, then verify against follower count and bio, exactly as on Instagram — and
treat a tiny follower count for a national operator as a failed lookup, not a finding,
until the search confirms it.

## Handle discovery

**Search for handles; never guess them.** Guessed handles land on impostor and personal
accounts — one guess in a live scan hit a personal account with 30 followers whose name
merely resembled the brand.

Verify each handle against follower count and bio before collecting. Be suspicious of a
follower count that looks wrong for the brand's real-world size; a major operator with
129 followers is probably a secondary or regional account, not the main one.

**Aggregator pages lie.** Search results and Linktree both gave one brand's handle as
`@conaturalintl`, which does not resolve; the live account was `@conatural`, found in the
brand's own website footer. **Check the brand's own site footer** before trusting a
third-party listing.

**A wrong handle fails silently.** Instagram serves the login wall, not a 404, so the
collector sees a page and writes a row with blank followers and zero posts. It now warns;
treat that warning as a hard stop, not a note.

**Absence is data.** Two substitutes in one scan had no Instagram at all. That is a
finding about where the category competes, not a hole in the sample.
