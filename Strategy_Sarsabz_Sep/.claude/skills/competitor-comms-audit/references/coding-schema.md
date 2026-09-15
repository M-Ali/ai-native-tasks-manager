# Coding schema

Every ad becomes one row. Coding consistently is what makes the counts mean anything, so read
this before coding and have everyone else coding read it too.

## Contents

- [The rule that resolves most disagreements](#the-rule-that-resolves-most-disagreements)
- [Fields](#fields)
- [Claim territory](#claim-territory) — the most important field
- [Emotional register](#emotional-register)
- [Proof device](#proof-device)
- [Other controlled fields](#other-controlled-fields)
- [Calibrating between coders](#calibrating-between-coders)
- [Adapting the vocabulary to a category](#adapting-the-vocabulary-to-a-category)

## The rule that resolves most disagreements

Code **what the ad asks the viewer to feel and believe**, not what the product technically is.
Two ads for the same policy can occupy completely different territories — one selling fear of
leaving a family unprovided for, the other selling the pride of a child's graduation. The
product is identical; the communications territory is not, and it is the territory we are
mapping.

When an ad genuinely does two things, put the dominant one in `claim_primary` and the
supporting one in `claim_secondary`. Dominant means: what the opening three seconds, the
headline, or the largest visual element commits to.

## Fields

| Field | Type | Notes |
|---|---|---|
| `ad_id` | auto | Sequential, assigned by the sheet |
| `brand` | dropdown | From the brand list |
| `ring` | auto | Filled from the brand list |
| `channel` | dropdown | Where the ad ran |
| `date_first_seen` | date | Ad library start date, or upload date |
| `date_last_seen` | date | Blank if still active |
| `format` | dropdown | Static, carousel, video, long-form, print, OOH, radio |
| `duration_sec` | number | Video only; blank otherwise |
| `language` | dropdown | Urdu, English, mixed, regional, or the category's equivalents |
| `product_line` | dropdown | Category-specific; set in the brand list config |
| `claim_primary` | dropdown | See below |
| `claim_secondary` | dropdown | Optional |
| `register` | dropdown | See below |
| `proof_device` | dropdown | See below |
| `cta_channel` | dropdown | How the viewer is asked to act |
| `audience_signal` | dropdown | Who the ad visibly addresses |
| `addresses_objection` | yes/no | Does it tackle the category's core objection head-on? |
| `views` | number | Where the platform reports it |
| `source_url` | text | **Required.** Rows without provenance are dropped |
| `coder` | dropdown | Initials — enables the drift check |
| `notes` | text | Anything the schema cannot hold |

## Claim territory

The default vocabulary below is written for financial services and protection categories.
Adapt it per category — see the last section — but keep the count between six and ten. Fewer
and everything collapses into one bucket; more and coders stop agreeing.

| Code | The ad is essentially saying | Watch for |
|---|---|---|
| `protection` | If something happens to you, your family is covered | The default claim in insurance; usually crowded |
| `savings_return` | Your money grows; this is an investment | Often blurs with protection — code on emphasis |
| `education_marriage` | Specific future obligations for your children | Culturally powerful in South Asia; often its own territory |
| `retirement` | Your own later life, income after work | Distinct from savings when the ad shows an older self |
| `faith_compliance` | This is permissible under your beliefs | Takaful/Shariah positioning; often the whole proposition |
| `national_trust` | Backed by the state, the nation, something larger | Rare and hard to claim — note who can credibly use it |
| `convenience_digital` | Easy, fast, do it on your phone | Rising in most categories; often masks a weak proposition |
| `price_affordability` | Cheap, accessible, low entry | Where challengers go when they have nothing else |
| `brand_corporate` | Who we are, our scale, our anniversary | No product claim; still occupies space and spend |

If an ad fits none of these, code `other` and describe it in `notes`. If `other` exceeds ~10%
of rows, the vocabulary needs adapting — stop and fix it rather than coding around it.

## Emotional register

What the ad wants the viewer to *feel*. Six options, deliberately few:

- `fear` — loss, risk, what happens if you do nothing
- `aspiration` — the better life this makes possible
- `duty` — obligation to family, responsibility, the right thing to do
- `reassurance` — calm, safety, we will be here
- `pride` — dignity, achievement, national or personal
- `humour` — comedy as the primary vehicle

Register is where categories reveal themselves. A category where every brand codes `aspiration`
has left `reassurance` and `duty` wide open, and that is a finding.

## Proof device

What the ad offers as evidence for its claim. Often the most diagnostic field in a low-trust
category:

`sovereign_guarantee`, `ratings_awards`, `testimonial`, `claim_statistics`, `celebrity`,
`expert_authority`, `heritage_scale`, `none`.

`none` is a legitimate and common code. A category where most ads offer no proof at all is a
category where proof is available as a differentiator.

## Other controlled fields

- **`cta_channel`** — `agent`, `bank_branch`, `digital_direct`, `call_centre`, `retail`, `none`.
  Reveals distribution strategy, which usually matters more than creative for a media
  recommendation.
- **`audience_signal`** — `mass`, `affluent`, `women`, `youth`, `diaspora`, `sme_corporate`,
  `rural`. Code what the casting, language and setting actually signal, not the target the
  brand would claim.
- **`addresses_objection`** — define the category's single core objection at scoping ("will
  they actually pay out?", "is this permissible?", "is my data safe?") and write it into the
  sheet's README so every coder applies the same test.

## Calibrating between coders

Before splitting the work: everyone codes the **same ten ads** independently, then compares.
Reconcile every disagreement and write the resolution into `notes` on the schema. Thirty
minutes here is the difference between data you can aggregate and data you cannot.

`analyze_audit.py` reports each coder's distribution across `claim_primary` and `register`. If
one coder's profile diverges sharply from the others, that is drift, not insight — recode their
rows before analysing.

## Adapting the vocabulary to a category

Edit the `Codes` sheet in the generated workbook; the dropdowns and the analyser both read from
it, so nothing else needs changing. Keep the principles:

- Six to ten claim territories, mutually exclusive, named as **what the ad says** rather than
  what the product is.
- Register stays at the six above; it is category-independent and comparability across pitches
  is worth more than local precision.
- Never add a code mid-collection without recoding what came before. A vocabulary that shifts
  halfway through produces counts that mean nothing.
