# Scan coding schema

One row per post. Nine coded dimensions, plus provenance and computed cadence.

Read this before coding anything, and make everyone else coding read it too.
Consistency between coders is what makes the counts mean something. Code a shared
sample of ten posts as a group, compare, reconcile, then split the work — half an
hour here saves a day of incomparable data.

## Contents

- [Columns](#columns)
- [1. product_focus](#1-product_focus)
- [2. audience_generation](#2-audience_generation)
- [3. content_type](#3-content_type)
- [4. claim_primary / claim_secondary](#4-claim_primary--claim_secondary)
- [5. register](#5-register)
- [6. proof_device](#6-proof_device)
- [7. who_is_in_frame](#7-who_is_in_frame)
- [8. production_value](#8-production_value)
- [9. ecosystem_push](#9-ecosystem_push)
- [Cadence](#cadence-computed-never-coded)
- [Calibration](#calibration)

## Columns

| Column | Filled by | Notes |
|---|---|---|
| `post_id` | script | shortcode |
| `brand`, `ring` | script | from handles.csv |
| `handle`, `source_url`, `image_file` | script | provenance — a row without `source_url` is dropped |
| `date` | script | post date where available, else collection date |
| `format` | script | static / carousel / video |
| `reposted_from` | script | creator handle if the brand reposted someone else |
| `product_focus` | **you** | free text, then normalised |
| `audience_generation` | **you** | vocabulary below |
| `content_type` | **you** | vocabulary below |
| `claim_primary`, `claim_secondary` | **you** | category-specific |
| `register` | **you** | vocabulary below |
| `proof_device` | **you** | vocabulary below |
| `who_is_in_frame` | **you** | vocabulary below |
| `production_value` | **you** | vocabulary below |
| `ecosystem_push` | **you** | yes / no |
| `campaign_platform` | **you** | the recurring line or hashtag, if any |
| `addresses_objection` | **you** | yes / no |
| `language_lead` | **you** | urdu_first / english_first / equal / other |
| `coder`, `notes` | **you** | quote the actual copy in notes |

## 1. product_focus

**Free text, then normalised at analysis.** Name the specific thing being pushed —
`Jubilee Active`, `Thrive`, `PRIMUS`, `Globewell`, `Hemayah Pension Fund` — not the
generic category. If it pushes the master brand only, write `brand`. If nothing
identifiable, `none`.

This is the dimension that catches **convergence**: several brands funnelling every
message into an owned app or ecosystem while their claims look different on a message
map. Convergence is usually the biggest strategic fact in a category and the easiest
one to miss.

Watch for a brand whose product_focus is a *platform* rather than a *product* — that is
a brand rebuilding itself around retention rather than acquisition, and it vacates
whatever territory it used to hold.

## 2. audience_generation

Inference from casting, styling, device, setting and language register. **It is a read
of who the ad is aimed at, never a fact about who buys.** Label it as inference wherever
it is presented.

| Code | Signals |
|---|---|
| `genz` | under ~28. Phone-first framing, creator/UGC aesthetic, slang, vertical video, gamified rewards |
| `geny` | ~28-43. Young family, first home, career, app convenience, aspirational-but-settled |
| `genx` | ~44-59. Teenage children, education and retirement horizon, established home |
| `boomer_plus` | 60+. Retirement, health, legacy, grandchildren |
| `mixed` | deliberately multi-generational — family groups spanning ages |
| `unclear` | no human signal, or genuinely unreadable |

Do not code `mixed` as a shrug. Use it when the ad *deliberately* casts across
generations; use `unclear` when you cannot tell.

## 3. content_type

The honesty test. A brand can look busy while publishing nothing that builds it.

| Code | What it is |
|---|---|
| `brand_building` | Advances a proposition or idea. Would still mean something in six months |
| `tactical_promo` | A discount, offer, voucher, limited-time reward |
| `product_feature` | Explains what a product does or how to use it |
| `corporate_pr` | Plaques, MOUs, ribbon-cuttings, conference panels, awards, executive quotes |
| `recruitment` | Hiring posts, job fairs, careers programmes |
| `calendar_topical` | Independence Day, Eid, religious and national dates, weather advisories |
| `csr` | Charity, screening camps, public-service messages |
| `ugc_repost` | Creator content reposted by the brand |

The **brand_building : tactical_promo : corporate_pr** ratio per brand is often the most
damaging chart you can show a client about themselves. Compute it always.

`ugc_repost` should also be caught automatically by the collector from the post URL —
keep both, and reconcile disagreements.

## 4. claim_primary / claim_secondary

**Category-specific. Rewrite this list per scan** — twenty minutes with the client's own
product language. The other eight dimensions stay fixed so scans stay comparable; this
one must fit the category or the territory map is meaningless.

Financial services / insurance default:

`protection` · `savings_return` · `education_marriage` · `retirement` ·
`faith_compliance` · `national_trust` · `convenience_digital` · `price_affordability` ·
`health` · `brand_corporate` · `other`

Rules: `claim_primary` is what the ad leads with, not everything it mentions. Use
`other` sparingly — a high `other` rate means the vocabulary does not fit the category
and should be rewritten, not worked around.

## 5. register

The emotional key. `fear` · `aspiration` · `duty` · `reassurance` · `pride` · `humour`

An almost-empty `fear` column in a protection category is worth noticing — categories
that could use fear and do not have usually decided it backfires, or have simply never
tested it.

## 6. proof_device

What backs the claim. This is where most categories are hollow.

`sovereign_guarantee` · `ratings_awards` · `testimonial` · `claim_statistics` ·
`celebrity` · `expert_authority` · `heritage_scale` · `none`

`none` is the most common code in most categories and the most important number in the
scan. A category where 70%+ of ads prove nothing is a category where a single credible
number can take a position.

Distinguish `claim_statistics` (a number about performance — claims paid, complaints
resolved, uptime) from `ratings_awards` (a third party's badge). The first is far
harder to copy.

## 7. who_is_in_frame

The fastest read of whether a brand talks to customers or about itself.

`customer` · `celebrity` · `executive` · `staff` · `expert` · `none_people`

A grid dominated by `executive` is a diagnosis, not a style choice. Report the
customer : executive ratio per brand.

## 8. production_value

Reveals budget, seriousness and whether a brand is actually investing.

| Code | What it looks like |
|---|---|
| `studio` | Commissioned shoot, art-directed, talent |
| `stock` | Licensed library imagery |
| `template_graphic` | In-house layout, brand template, type on colour |
| `event_photo` | Photographs from an actual event |
| `ugc` | Creator-shot, phone-native |
| `ai_generated` | Synthetic imagery — increasingly common; note tells (hands, text, texture) |
| `unclear` | cannot tell |

A brand running `template_graphic` for twelve straight posts is not running a campaign,
whatever its captions claim.

## 9. ecosystem_push

`yes` / `no`. Does the ad drive to an owned app, portal or platform, as opposed to a
sale, an agent, a branch, or nothing?

Cross-tabbed with `product_focus`, this is the convergence measure: how much of a
category's output exists to drive app installs rather than to sell the category's
actual product.

## Cadence — computed, never coded

`analyze_scan.py` derives, per brand:

- **posts per week** across the collected window
- **days since last post**
- **active / moderate / slow** banding

Never eyeball "they post a lot". And state the window: with a capped sample this is a
rate over the recent period, not a long-run average.

**Pinned posts distort this badly.** Instagram serves up to three pinned posts at the top
of a grid, and a pin can be months older than everything else. Taken naively, a brand that
published five posts in six days reads as 0.6 posts/week. The collector flags the first
three grid positions as `pinned_guess` and `analyze_scan.py` excludes them from the span —
but check the flag against the dates before quoting a cadence number.

## Calibration

Before splitting work, everyone codes the same ten posts independently, then compares.
The dimensions that drift most are `content_type` (brand_building vs product_feature)
and `audience_generation`. Agree the boundary on those two explicitly.

`analyze_scan.py` includes a coder-drift check. Run it; if two coders diverge sharply on
one dimension, recode that dimension rather than shipping mixed data.
