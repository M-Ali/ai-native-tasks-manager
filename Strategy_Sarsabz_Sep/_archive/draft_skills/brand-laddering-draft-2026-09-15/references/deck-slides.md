# The four ladder slides, and how to render them

Every slide carries a footnote naming the framework, the sample and the file the numbers came from. A ladder
slide without its n is an opinion in a table.

## Slide 1 — the framework

A divider. Title: **The Brand Meaning Ladder**. One line: *a simple way to understand how consumers see a brand —
from what it is and does, to what it means to them*, then *from functional benefits → to emotions → to values and
purpose*. Nothing else; the work is on the slides after it.

## Slide 2 — placement

**Headline pattern:** *Levels of branding: where the category is, and where [CLIENT] is talking.*
**Standfirst:** *Placement computed from the N coded posts — each brand's dominant claim mapped to a level, not asserted.*

Six rungs, highest at the top. Each rung carries the level name, its phrase ("Reason to buy me"), and the brands
that sit there with their computed means. Empty rungs are stated, not hidden: *"Nobody in this category."*

A right-hand marker calls out up to four rungs:

| Marker | On the rung where |
|---|---|
| `WHERE [CLIENT] IS TALKING` | its own coded posts cluster |
| `THE DIFFERENTIATION GAP` | the level between where it talks and where it must earn |
| `WHERE [CLIENT] MUST EARN` | the recommended placement |
| `WHERE THE CATEGORY SITS` | the modal level across competitors |

**Conclusion bar:** the gap in one sentence — e.g. *"[CLIENT] is communicating roughly two levels above where it is
delivering — and the ladder does not let a brand skip."*

This slide renders directly from `category-creative-scan`'s `build_deck.py` via a `_ladder` block in `reads.yaml`:

```yaml
_ladder:
  headline: "Levels of branding: where the category is, and where Sarsabz is talking"
  standfirst: "Placement computed from the 210 coded posts - each brand's dominant claim mapped to a level, not asserted."
  levels:                                  # highest level first; up to 6
    - name: "6. Mission / philosophy"
      says: "I change things you would like to"
      who: "Nobody in this category."
      mark: ""                             # optional right-hand marker
    - name: "5. Human values / way of life"
      says: "I think like you do"
      who: "Sarsabz - Salam Kissan, Tabeer, in the pride register"
      mark: "WHERE SARSABZ IS TALKING"
  conclusion: "One sentence naming the gap."
  source: "Framework: Levels of Branding. Placement is each brand's mean level across its coded posts; a claim map, not a quality judgement."
```

## Slide 3 — the factor inventory

**Headline:** *The Brand Meaning Ladder for [category]: the factors at each level.*
**Standfirst:** *Current = seen in the category's N posts. Potential = relevant to [category], not yet claimed.
Unsupported = little link to [category]. † = added for this category.*

Columns: **LEVEL** · **CURRENT** · **POTENTIAL** · **UNSUPPORTED**. Each level row gives the level name and its
diagnostic question ("What do you have?").

**Conclusion bar:** where the current factors cluster, and the standing caveat — *higher is not better: a factor is
only worth owning if it is relevant, credible, distinctive and defensible.*

## Slide 4 — the client's ladder and the verdict

The slide the recommendation lives on. Two versions, depending on whether customer evidence exists.

**With customer evidence** — columns: **LEVEL** · **WHAT BRANDS CLAIM** · **WHAT CUSTOMERS JUDGE ON** · **OPEN — NOBODY CLAIMS**.

**For the client specifically** — columns: **LEVEL** · **FACTORS IN THE CATEGORY** (split *In use* / *Open*) ·
**[CLIENT] TODAY** (split *Says:* with its post count / *Heard:* with the customer evidence) · **PLACEMENT** (the
verdict word plus a three-to-five-word gloss, e.g. `PROVE — publish terms and speed`).

**Standfirst** states both samples: *[CLIENT] says = its N coded posts, classified by level (mode L1, mean 2.4).
Consumers say = N comments. Higher is not better — the target is the deepest level [CLIENT] can prove.*

**Conclusion bar:** the placement sentence — *place [CLIENT] at Level N: [territory]. Nobody claims it, it answers
[the customer evidence], and [CLIENT] can prove it with [commitment]. Level N+1 is earned from there; [existing
asset] stays a Level 1 asset.*

## Rendering

Slides 3 and 4 are dense tables. Build them with `python-pptx` as text boxes on a blank layout (the pattern in
`build_deck.py`): one row per level, ~0.78in high, 9pt body, the level column ~1.7in wide. Keep every column's
text to the width given — a ladder table that reflows is unreadable in a pitch room.
