---
name: category-creative-scan
description: "Scan a whole category's live social creative brand by brand - collect every competitor's recent posts and the actual artwork, code them on a nine-dimension schema (product focus, audience generation, brand-building vs tactical, cadence, and more), and produce a brand-by-brand deck plus a gap analysis. Use when the user wants to scan a category, see what competitors are posting, compare brands' social output, find the white space, or asks for a competitive creative review, category scan, social audit, or brand-wise deck. Triggers on: scan the category, what are competitors posting, category scan, creative audit, brand-wise deck, competitor social, what's the gap."
---

# Category creative scan

Collect what a category is actually publishing, look at the artwork, and find the gap.

This is the **collection and pattern** half of competitive work. Its sibling skill
`competitor-comms-audit` carries the strategic method — rings, message territory,
share of voice, how to write the finding. Use them together: scan first, then write
the finding. If you only need the argument and already have the ads, go straight to
`competitor-comms-audit`.

## Why this exists as its own skill

Category scans recur constantly and the manual version is slow and inconsistent. The
expensive parts are collection (access keeps changing) and coding (drifts between
people). Both are solved here. What is never automated is the read — that is yours.

## The nine dimensions

A message-territory map alone tells you what brands *say*. It misses how they behave,
which is where most gaps actually are. Code every ad on all nine. Full vocabularies in
`references/scan-schema.md`.

| # | Dimension | The question it answers |
|---|---|---|
| 1 | `product_focus` | What product, platform or ecosystem is this brand converging on? |
| 2 | `audience_generation` | Who is this cast, styled and written for — GenZ, GenY, GenX, older? |
| 3 | `content_type` | Brand-building, tactical promotion, corporate PR, recruitment, calendar, CSR? |
| 4 | `claim_primary` | What is it claiming? |
| 5 | `register` | In what emotional key? |
| 6 | `proof_device` | What backs the claim up — or nothing? |
| 7 | `who_is_in_frame` | Customer, celebrity, executive, staff, or no people? |
| 8 | `production_value` | Studio, stock, UGC, AI-generated, event photo, template graphic? |
| 9 | `ecosystem_push` | Does it drive to an owned app or platform? |

Cadence — **posts per week, and days since the last post** — is computed from dates, not
coded. Never eyeball "they post a lot."

### Why these four earn their place

**`product_focus`** catches convergence. When three brands in a category all funnel every
message into a rewards app, that is the category's centre of gravity, and it is usually
invisible on a claim map because the claims differ while the destination is identical.

**`audience_generation`** catches abandonment. Categories drift young — the casting, the
platforms, the tone all follow — and the segment with the money quietly stops being
addressed. Read it off the creative: who is cast, what they wear, the device in frame,
the language register. State it as inference, never as fact about who actually buys.

**`content_type`** is the honesty test. A brand can look busy and be publishing nothing
but plaques and hiring posts. The brand-building : tactical : PR ratio is often the single
most damaging chart you can show a client about themselves.

**`who_is_in_frame`** is the fastest read of whether a brand is talking to customers or
about itself. A grid full of executives receiving awards is a diagnosis.

## Workflow

### 1. Frame the category

Write down, before collecting: the category, the client, the window, and the competitive
set in four rings — client / direct / adjacent / substitute. Do not default to the
head-to-head list; substitutes and non-consumption usually take more volume than named
rivals. `competitor-comms-audit` has the ring method in full.

### 2. Check where the audience actually is, before choosing a platform

Do not assume the category lives on Instagram because the collector does.

```bash
# candidates.csv: brand,platform,url - one candidate page per brand per platform
python scripts/sweep_presence.py candidates.csv --out <workspace>/scan/presence.csv
```

Probes each candidate page logged-out and records follower/subscriber count, page name
and a bio snippet. Two browser calls per URL, because Facebook and TikTok both serve an
interstitial on first load — see `references/access.md`.

This costs twenty minutes and it has changed the framing of a scan. In Pakistani fixed
broadband, **every brand was an order of magnitude larger on Facebook than on Instagram**:

| | Instagram | Facebook | |
|---|---:|---:|---|
| PTCL | 91K | 1.3M | 14× |
| Nayatel | 19K | 1.1M | **57×** |
| Transworld Home | 13K | 381K | 30× |
| Wateen | 1.2K | 93K | 76× |

Two things fall out of that, and neither is recoverable later:

- **A brand can look absent when it is merely elsewhere.** Fiberlink has no Instagram at
  all and 23K followers on Facebook. Recorded as "no account", that is simply wrong.
- **A brand's weight can be misread.** Nayatel reads as a mid-size regional player on
  Instagram and is comparable to the market leader on Facebook.

Sweep first, then say *in section 1 of the output* which platform you read and what
multiple of the audience it represents. Scanning the minor channel is a legitimate
choice — the images are there, and it is where paid-looking creative tends to be
published — but it is only legitimate when stated.

**Absence is a finding here too, and it is cheap to establish.** In the same scan,
TikTok returned "Couldn't find this account" for the client and its nearest rival, and
eight searches surfaced no official account for any brand in the category. That settles
whether to build a collector for it.

**Do not publish a `suspect` or `no_count` row.** A tiny follower count for a national
operator is a failed lookup, not a finding; a `no_count` is usually a wrong candidate
URL, not an absence. In one sweep a YouTube probe resolved to a channel called "Rusty
Stainless". Search the handle properly or leave the cell empty.

### 3. Resolve handles

```bash
python scripts/make_handles.py brands.csv --out <workspace>/scan
```

Then fill the `handle` column by search — **search, never guess**. Guessed handles land
on impostor and personal accounts. Verify each against follower count and bio.

**A brand with no account is a finding, not a gap.** Leave the handle blank and record
why; the scripts carry absences through to the output.

### 4. Collect

```bash
python scripts/collect_ig.py <workspace>/scan/handles.csv --out <workspace>/scan --depth 24 --enrich
```

`--enrich` opens each post for its **true date and full caption**, which the profile grid
does not carry. It costs one extra page load per post and is worth it: without real dates
cadence is guesswork, and without captions you are coding from images alone.

Headless Chromium via the `browser-automation` skill's `browser.mjs`. No login, no cookie,
no API key. Writes `profiles.csv`, `posts.csv`, and every image into `media/<handle>/`.

**Known ceiling: 24 posts per profile.** The grid serves 12, then a "Show more posts"
button yields one further tranche of 12 without a login; after that it disappears. Pass
`--depth 24` — it costs one extra click per brand and doubles the read, which matters
because a 12-post window is barely two weeks for an active brand. Beyond 24 needs an
authenticated session.

Twenty-four recent posts is a sound current-period read; it is not a history. Consequences
you must honour:

- Equal N per brand means **share of voice cannot be measured**. Do not present one.
- Cadence from 12 posts is a rate over a short window. State the window.
- For depth, an authenticated session is required — see `references/access.md`.

The collector **warns loudly when a handle returns no posts or followers**. Do not ignore
it: a wrong-but-plausible handle gets served the login wall, not an error, and would
otherwise land in `profiles.csv` as a normal row and ship as a finding. Use `--append`
when re-collecting a single brand, or you will truncate the whole file.

### 5. Seed the capture sheet

```bash
python scripts/seed_capture.py <workspace>/scan
```

Carries every collected post into `capture.xlsx` with its provenance — post id, brand,
date, source URL, image file, and the caption in `notes` so you can read it while coding.
Re-runnable: existing codes are matched on `post_id` and preserved, so collecting more
brands later does not lose work.

### 6. Look at the creative

```bash
python scripts/contact_sheets.py <workspace>/scan
```

Then **actually read the sheets**. This is the step that cannot be skipped or delegated
to a caption. Captions describe intent; artwork shows execution, budget, casting and
production value. Half the findings in a good scan are visible only in the image.

Code what you can read. Anything illegible at thumbnail size gets `unclear` — open the
full-size file, or leave it `unclear` and say so. Never infer a code from a brand's
general reputation.

### 7. Code, then apply

Code in whatever tool suits you, then land it safely:

```bash
python scripts/apply_codes.py <workspace>/scan/codes.csv --workspace <workspace>/scan
```

`codes.csv` needs `post_id` plus whichever columns you are setting. Every value is
validated against the vocabulary **before** anything is written — a typo fails with a
list of what was wrong rather than silently poisoning the analysis. A bad code is worse
than a blank one, because it looks like data.

Coding 150+ posts inline in a script hits shell argument limits on Windows. Use the CSV.

### 8. Analyse

```bash
python scripts/analyze_scan.py <workspace>/scan/capture.xlsx --out <workspace>/scan
```

Produces `scan_findings.md` plus CSVs: product-focus map, generation map, content-type
balance per brand, cadence table, ecosystem convergence, who-is-in-frame ratio, claim ×
register territory map, and proof/objection rates.

### 9. Build the brand-wise deck

Author `reads.yaml` — one entry per brand: a headline read, three to five specific things
they are saying **with real copy quoted**, and a strategic note. Then:

```bash
python scripts/build_deck.py <workspace>/scan --title "..." --client "..."
```

Each brand slide pairs that brand's contact sheet with the read and a counted footer, so
every judgement on the page is anchored to a number from the sheet.

### 10. Find the *why*, not just the *what*

A pattern is an observation. The mechanism behind it is the insight, and it is what a
client is actually paying for.

When the scan shows every brand doing the same thing, ask what structural fact makes it
rational. The answer is usually in the market's plumbing — distribution, regulation,
channel economics, penetration — and it is usually **documented somewhere public**. Search
the competition regulator, the industry body, the trade press. A sourced mechanism beats a
plausible story, and it is the difference between "they are all building apps" and "they
are all building apps because none of them owns their customer."

Two tests before it goes on a slide:

- **Can you cite it?** If the mechanism rests on your own reasoning alone, label it as a
  hypothesis, not a finding.
- **Does it change the recommendation?** A mechanism that leaves the advice unchanged is
  trivia. The useful ones tell the client which competitor behaviours to ignore, because
  they solve a problem the client does not have.

Put a `why` line on every brand slide too. "EFU runs three ecosystems" is a fact; "three
ecosystems at once is what bancassurance dependence looks like from the inside" is a read.

### 11. Write the finding

One sentence naming the vacant territory; one paragraph proving it. Hand off to
`competitor-comms-audit` §6 for the structure and the defensibility test.

**Pair the supply side with the demand side wherever you can.** A scan reads what brands
publish, which is the half the client already knows. Run customer comment or review data
through `comment-analysis` and put the two next to each other: the gap between what a brand
says and what its customers report is the finding a client cannot argue with. In one scan
the client's own creative claimed national identity while its customers were describing the
company in the language of fraud - neither dataset showed that alone.

**A staged model turns "move up" into a specific claim.** Placing every competitor on a
ladder from the coded claims - rather than asserting where they sit - converts a vague
ambition into an argument with arithmetic behind it: this is where the category is, this is
where you are talking, and this is the rung you have not earned. `_ladder` renders it.

## Slides the deck can build

`build_deck.py` renders these from `reads.yaml` when the keys are present. All are
optional; the brand slides work alone.

| Key | Slide |
|---|---|
| `_category` | commentary under the four computed category statistics |
| `_dimensions` | commentary under the three behavioural rankings |
| `_why` | **the mechanism** — headline, standfirst, up to four sourced drivers, a conclusion, a source line |
| `_audience` | who the work is cast for, computed, with the inference caveat |
| `_grid` | a plain matrix - brands down the side, whatever you measured across the top; unverified cells print `n/v`, never blank |
| `_voice` | **the demand side** - up to six customer verbatims with notes, a conclusion and a source line |
| `_ladder` | place the category on a staged model (levels of branding, a maturity model) and mark where the client is versus where it should go |
| `_readings` | three ranked panels for findings that do not fit elsewhere |
| `why:` on a brand | one red line under that brand's note — the mechanism, not the observation |
| `_finding` | the closing argument |
| `_own` | what the client can own: a number, why it beats the alternative, three pillars |

## Rules that keep this honest

- **Sweep presence before you pick a platform, and state the one you read.** The category
  is often an order of magnitude bigger somewhere else, and a brand that is merely
  elsewhere will otherwise be recorded as absent.
- **Every row needs a `source_url`.** Rows without provenance are dropped at analysis.
- **Code from the artwork, not the caption.** Note when you could not.
- **`unclear` is a valid code.** A guessed code is worse than a gap.
- **Never present share of voice from a capped sample.**
- **Organic is not paid.** Everything here is owned-channel output. It says nothing about
  spend, reach or media weight. Say so in section 1.
- **Reposted creator content is not the brand's own creative.** The collector flags it;
  keep the flag, and count it separately — a feed padded with UGC is a finding.
- **Generation and audience codes are inference from casting.** Label them as such.
- **Never count your own defaults and call it a finding.** Any coded field with a default
  value measures the coder, not the category. `analyze_scan.py` reports script usage from
  the raw caption text for exactly this reason — a `language_lead` tally would have said
  97% English when the coder had simply defaulted to it. Where a finding could be an
  artefact of your coding, re-measure it independently before it goes near a slide.
- **Credit small brands honestly.** The sharpest work in a category often comes from the
  smallest advertiser. A scan that only flatters the leaders is not a scan.

## Adapting to a category

The claim vocabulary in `references/scan-schema.md` ships tuned for financial services
and **must** be rewritten per category — twenty minutes with the client's own product
language. The other eight dimensions are category-agnostic and should not be changed;
their comparability across scans is the point.
