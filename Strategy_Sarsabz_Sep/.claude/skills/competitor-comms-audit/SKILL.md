---
name: competitor-comms-audit
description: Run a competitor communications audit for a pitch, tender or brand review - collect competitors' live advertising, code every ad on a fixed schema, and produce a message-territory map, share-of-voice-vs-share-of-market comparison and a vacant-territory finding. Use this whenever the user mentions competitive analysis, competitor analysis, category review, share of voice, message territory, ad audit, competitive landscape, "what are competitors saying", or is preparing a pitch, RFP response, tender submission or new-business deck that needs a competitive section - even if they only say "look at the competition" or "analyse the category".
---

# Competitor communications audit

A method for answering one question with evidence: **which message territory is unoccupied,
and can our client credibly own it?**

This is a communications audit, not market research. The client's own commercial team knows
their market share better than we ever will. What an agency is uniquely able to tell them is
what the category *sounds like* — who is claiming what, to whom, in which channels, at what
weight — and where the gap is.

## The test this work must pass

If the analysis could be deleted without weakening the campaign concept that follows, it has
failed. The audit exists to make the concept look inevitable rather than invented. Write it as
one argument with the concept, not as a standalone research report.

Two failure modes to avoid:

- **The encyclopedia.** Forty pages of category statistics with no point of view. Reads as
  homework. Evaluators score it low because any agency could have produced it.
- **The assertion.** A confident claim about "what the category is doing" with no sample frame
  behind it. If you cannot say how many ads you looked at, over what period, from which
  sources, you are guessing.

## Workflow

### 1. Scope it — do this before collecting anything

Establish, and write down:

- **The brief.** What product, audience and objective is the campaign actually for? If the
  client has not said, ask them. On a tender there is usually a formal clarification window —
  use it, and check the procurement portal for a scope or lots section that may not be in the
  main document.
- **The competitive set, in rings.** Do not default to the obvious head-to-head list. Most
  categories lose more volume to non-consumption and informal substitutes than to named rivals,
  and market leaders in particular are badly served by an "us vs them" frame. Build:
  - Ring 0 — the client
  - Ring 1 — direct competitors
  - Ring 2 — adjacent/faith-based/format variants competing for the same need
  - Ring 3 — substitutes and non-consumption (including advertisers outside the category who
    compete for the same wallet or the same job-to-be-done)
- **The window.** Usually the last 12 months. State it and stick to it.
- **The time budget.** Collection expands to fill whatever it is given. Decide up front how
  many days this gets, and protect the creative work that depends on it.

### 2. Generate the capture sheet

```bash
python scripts/make_capture_sheet.py brands.csv --out <workspace>/audit --category "<category name>"
```

`brands.csv` needs `brand,ring,notes`. The script writes a workbook with locked controlled
vocabularies as dropdowns, so several people can code in parallel without drifting apart.

Read `references/coding-schema.md` before coding anything, and make everyone coding read it
too. Consistency between coders is what makes the counts mean something.

### 3. Collect

Read `references/sources.md` for where to look per channel and the gotchas in each.

The spine is whichever ad-transparency library covers the client's main paid channels — it is
dated, complete for active ads and verifiable, which matters when an evaluator asks where a
number came from.

**Parallelising:** with more than about six brands, fan out subagents — one per brand, each
returning rows in the capture schema — rather than collecting serially in your own context.
Give each agent the coding schema and the exact column list, and have it report rows only, no
commentary. Then paste the rows into the sheet and code centrally, or spot-check a sample if
the agents coded directly.

Collect **the ad, not the impression of the ad**: every row needs a URL or screenshot
reference. Rows without provenance get dropped at analysis.

### 4. Code every ad on the fixed schema

The schema is in `references/coding-schema.md`. It captures, per ad: claim territory, emotional
register, proof device, call-to-action channel, audience signal, product line, format,
language, and whether the ad addresses the category's core objection.

This coding *is* the analysis. Everything downstream is presentation. Code a shared sample of
ten ads as a group first, compare, and reconcile disagreements before splitting the work — half
an hour here saves a day of incomparable data.

### 5. Analyse

```bash
python scripts/analyze_audit.py <workspace>/audit/capture.xlsx --out <workspace>/audit
```

Produces a findings report with: claim × register crosstab (the territory map data), per-brand
claim profiles, channel and language splits, share-of-voice proxy against share of market,
objection-handling rates, and a coder-drift check.

**On share of voice:** unless you have bought monitoring data, you cannot measure spend. Count
ads, and where available creative volume and view counts, and label it a *proxy for presence*.
Never present an estimated spend figure as if it were measured — one challenged number
discredits the whole section.

**Share of voice against share of market** is usually the single most valuable chart. It shows
whether a brand's voice matches its actual position. A leader being out-talked by challengers a
fraction of its size is a strategic finding an evaluator can act on.

### 6. Write the finding

The output is not the data. It is one sentence naming the vacant territory and one paragraph
proving it, supported by:

- **The territory map** — brands plotted on claim × emotional register, showing crowding and
  white space.
- **Share of voice vs share of market.**
- **The claim-to-reality gap** — what competitors promise in advertising versus what the
  category actually delivers. This is usually where the wedge is.

Then state what the client can credibly own — credibly being the operative word. Anchor it in
an asset competitors structurally cannot copy (ownership structure, scale, heritage,
distribution, guarantee), not in a tone of voice anyone could adopt next quarter.

Where a territory is commercially or politically sensitive for the client, present it as a
choice with its trade-offs rather than resolving it for them. Clients reward being given an
informed decision; they resent being handed a fait accompli.

### 7. Keep it short

A competitive section is judged on sharpness, not weight. Ten to fifteen pages of evidence with
a point of view beats forty pages of category survey — and it is the version that can actually
be finished on a pitch timeline.

## Output structure

Use this shape unless the client has specified their own:

```
1. What we looked at        - sample frame: brands, window, channels, ad count
2. How the category talks   - territory map, claim profiles, the crowding
3. Who is heard             - share of voice vs share of market
4. What nobody is saying    - the vacant territory, and the evidence for it
5. What [client] can own    - the defensible asset, and why competitors cannot follow
```

Section 1 is not throat-clearing. Stating the sample frame up front is what separates this from
opinion, and evaluators notice.
