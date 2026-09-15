---
name: brand-laddering
description: "Place a brand and its category on the Brand Meaning Ladder - the six levels from product identification, through attributes, benefits and psychological associations, to values and purpose - from coded advertising and customer evidence rather than assertion, then decide the deepest level the brand can prove and stand on. Use whenever the user asks where a brand should sit, whether it should move up, how to go from functional to emotional, or what the brand stands for. Use it after a category creative scan or comment analysis to turn coded posts into a placement, and before writing a campaign concept or positioning statement, so the idea is pitched at a level the brand can defend. Triggers on: brand ladder, laddering, levels of branding, brand meaning, means-end, brand pyramid, where should the brand sit, move up the ladder, functional to emotional, purpose-led positioning, brand depth, positioning level."
---

# Brand laddering

Decide the deepest level of meaning a brand can **prove**, and show the category's own levels next to it,
so "move up the ladder" becomes a placement with arithmetic behind it instead of a slogan.

The ladder is a claim map, not a quality ranking. **Higher is not better.** A brand talking two levels above
what it delivers is not aspirational, it is unbelievable — and that gap is usually the most useful finding
the work produces.

## The six levels

Read `references/levels.md` for the full definitions, the diagnostic question at each level, and the factor
pools to populate them.

| # | Level | What the brand is saying | The customer's question |
|---|---|---|---|
| 1 | Product identification | "I am here" | Who are you? |
| 2 | Product attributes | "Reason to buy me" | What do you have? |
| 3 | Product benefits | "Here is what I can do for you" | What does it do for me? |
| 4 | Psychological associations | "I am like you; you are like me" | What does choosing you say about me? |
| 5 | Human values / way of life | "I think like you do" | What do you believe? |
| 6 | Mission / philosophy | "I change things you would like to" | What change do you want to create? |

## What this skill needs before it can run

It is a *reading* of evidence, so it needs evidence:

- **Coded competitor advertising** — the output of `category-creative-scan` (`codes.csv` / `capture.xlsx`),
  one row per ad with a dominant claim. Without it, placement is assertion and the slide is worthless.
- **Customer evidence** where it exists — `comment-analysis` output, reviews, research. This is what turns a
  claim map into a judgement about what a brand can *prove*.
- **The client's own output**, coded the same way, so "where they are talking" is measured, not remembered.

If the user has none of this, say so and offer to run the scan first. A ladder built from memory is the exact
failure this skill exists to prevent.

## Workflow

### 1. Map each claim in the coding vocabulary to a level

Write the mapping down before computing anything — it is the one judgement everything else rests on, and it
must be visible for challenge. Start from `assets/claim_to_level.example.yaml`.

Two rules that keep the mapping honest:

- **Map on what the ad asks the viewer to believe, not what the product is.** A fertilizer ad that says "10%
  more yield" is Level 3 (benefit). The same product shown as "the mark of a progressive farmer" is Level 4.
- **A claim can appear at more than one level in different executions.** Map the *dominant* claim per row and
  say in the footnote that counts show where factors appear, not per-level totals.

### 2. Compute placement — never assert it

```bash
python scripts/place_brands.py <codes.csv> --map <claim_to_level.yaml> --out <workspace>/ladder
```

Writes `brand_levels.csv` (per brand: n, mean level, modal level, distribution), `level_mix.csv` (the
category's spread) and `placement.md`. It **fails loudly on any claim value not in the map** rather than
silently dropping rows — an unmapped claim is a hole in the analysis, not a rounding error.

Report the mean to one decimal and always with n. A brand with 24 coded posts at mean 2.4 is a reading; a
brand with 3 posts is an anecdote, and the script flags it.

### 3. Build the factor inventory for the category

For each level, list the factors **currently claimed** (from the coded set, with counts), those **open** — 
relevant and credible in this category but claimed by nobody — and those **unsupported**, which belong to other
categories and would not be believed here. Mark any factor you added for this category with `†` and say so in
the footnote.

The unsupported column matters more than it looks: it is what stops a purpose slide drifting into borrowed
language ("making healthcare accessible" in a fertilizer deck).

### 4. Put the demand side next to the supply side

Three columns: **what brands claim** · **what customers judge on** · **what nobody claims**. Customer evidence
usually sits one level *below* where brands are talking, and often names a value by its absence — in one
telecoms study, fairness appeared only as "fraud" and "zulm". That mismatch is the finding.

### 5. Decide the placement, level by level

For the client, at every level, record three things: what it **says** (count of its own coded posts), what is
**heard** (customer evidence), and a **verdict**. Use a fixed verdict vocabulary so the recommendation is
unambiguous:

| Verdict | Means |
|---|---|
| `HOLD` | Already owned; keep it, do not spend against it |
| `PROVE` | Claimed but not believed; needs evidence before anything above it works |
| `PLACE HERE` | The recommended position — the deepest level the brand can prove now |
| `NEXT` | Earned from the placement once proof lands |
| `LATER` | Credible only after the level below is established |
| `NOT NOW` | No permission; claiming it would invite ridicule |

**Exactly one `PLACE HERE`.** If two levels are candidates, the lower one is the answer — the ladder does not
let a brand skip a rung, and a brand that has not proved Level 3 cannot rent Level 5 with a film.

### 6. Write the conclusion as a sentence a client can act on

The pattern that works: *place [brand] at Level N: [the territory, in the customer's words]. Nobody in the
category claims it, it answers [the customer evidence], and [brand] can prove it with [the specific,
checkable commitment].*

Then say what happens to the asset the brand is attached to today. Heritage usually stays a Level 1 asset:
real, useful for recognition, and not a differentiator when 13 of 14 brands claim it.

## Slides this produces

`references/deck-slides.md` has the four slide specifications and the `_ladder` YAML block that
`category-creative-scan`'s `build_deck.py` renders directly.

1. **The framework** — what the ladder is, one line per level.
2. **Placement** — every brand on its rung, from computed means, with the client's position marked.
3. **The factor inventory** — level × (current / potential / unsupported).
4. **The client's ladder** — level × (factors in category / what it says / what is heard / verdict).

## Rules that keep this defensible

- **Placement is computed from coded evidence and says so on the slide**, with n and the source file.
- **Higher is not better.** Say it on the slide. The target is the deepest level the brand can prove.
- **No skipping rungs.** A recommendation to jump from attributes to purpose is a recommendation to be
  disbelieved.
- **Customer evidence outranks brand claims** when the two disagree. What is *heard* decides what can be claimed.
- **A level with no evidence is "not established", not "an opportunity"** — unless you can name the proof that
  would establish it.
- **Label the samples.** Coded ads are a supply-side read; comment corpora are self-selected. Neither is research.
- **Give the framework its name and its limits.** It is a way of organising claims, not a law of branding.

## Worked example

`references/levels.md` carries a full worked example from a Pakistani fixed-broadband study: 336 coded posts
across 14 brands, the category clustered at Levels 1-3, the client talking at Level 5 (heritage and nation, in
the pride register) while its customers judged it at Level 3 and below. The recommendation placed it at Level 3
— "fair terms you can count on" — with Level 4 as `NEXT` and purpose as `NOT NOW`. That is the shape of a good
output: uncomfortable, specific, and derived from counts the client can check.
