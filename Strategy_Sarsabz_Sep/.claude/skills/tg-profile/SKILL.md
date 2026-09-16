---
name: tg-profile
description: "Build a target-group profile card - one audience, one slide: who they are, their behavioural mindset, their pain points, their buying triggers, their interests with affinity indices, and where their media time goes. Use whenever the user asks for a TG profile, target group, target audience profile, audience persona, consumer persona, customer profile, buyer persona, segment profile, 'who are we talking to', an audience card or persona slide for a pitch or media plan, or wants interest affinity indices computed from an analytics interests export. Also use when a strategy, campaign concept or media plan needs the audience made concrete before creative work starts. Triggers on: TG profile, target group, target audience, persona, audience profile, consumer profile, buyer persona, segment, affinity index, media consumption, who is the audience."
---

# TG profile

A target-group profile turns "male farmers 25-55" into a person a creative team can picture and a media planner can buy
against. It fails in two opposite ways: as a stock-photo fiction nobody can act on, or as a table of demographics with no
human being in it. This skill produces one slide that holds both, where **every line traces to evidence the client can check.**

## The rule that makes it worth reading

**A persona is an argument about a real audience, not a character sketch.** So each block on the card carries its evidence:
a count, a verbatim, a dataset, a page of the brief. The card has room for about six pain points; if you cannot evidence
six, show four and say what would be needed for the rest. An unevidenced line is worse than a missing one, because the
client will act on it.

Three habits keep this honest:

- **Name the population behind every number.** An interest index computed from a news site's web analytics describes that
  site's visitors. It does not describe farmers, shoppers or "Pakistanis", and a card that implies otherwise is wrong even
  when the arithmetic is right.
- **Never invent a media percentage.** "TV 65%" with no survey behind it is the single most-copied fiction in persona decks.
  Either cite the survey, leave the number out and describe the channel qualitatively, or leave a labelled blank for the
  client to fill.
- **Pain points come from what the audience said, not from what the brand fears.** Run `comment-analysis` (or use its output)
  before writing this section. The register matters: "it's expensive" and "they take money from the mafia to make videos" are
  different briefs.

## Where this sits

`comment-analysis` (what they say) → **TG profile** (who they are) → `communication-strategy` (what to tell them) →
`big-idea` → `campaign-concept`. The profile is what makes the strategy's audience section concrete, so build it before
the message architecture, not after.

## The eight blocks

| Block | What it answers | Evidence it needs |
|---|---|---|
| Identity | Name, role, age, location | The brief's target definition, with page |
| Household | Family status, home, land or assets | The brief, client data, or census; say which |
| Behavioural mindset | How they decide, what they fear | The comment corpus' register, not adjectives |
| Pain points | What blocks them, in their words | Counts and verbatims from the corpus |
| Buying triggers | What moves them to act | Corpus evidence, seasonality from the brief |
| Interests + affinity index | What else they are into | An analytics export, with its population named |
| Media consumption | Where to reach them | A media survey, or evidenced channel behaviour |
| Sources | Everything above, traceable | File names, pages, dataset dates |

Not every audience has all eight. A rural audience usually has no analytics interest data, and saying so on the card is
better than borrowing an urban dataset. **A block you cannot evidence is replaced by one line naming what would fill it.**

## Workflow

### 1. Define the segment before you touch data

Write the segment in one line from the brief: age, gender, geography, income or land, and any behaviour the brief names
("TikTok-engaged", "12-150+ acres"). If the brief gives two definitions, put both on the card and flag the conflict — that
conflict is usually worth more to the client than the card itself.

### 2. Compute interest affinity indices, if a suitable export exists

```bash
python scripts/interest_index.py <interests.csv> --gender male --age 25-34,35-44,45-54 --out <dir>
```

The script reads the wide "Interests free-form" export (two header rows: gender, then age), computes for each interest

```
index = (segment share of that interest) / (all-users share of that interest) x 100
```

and writes `interest_index.csv` plus a printed top list by index and by volume. An index of 137 means the segment is 37%
more likely than the whole file to carry that interest.

Read the warnings it prints:

- **Small bases lie.** An interest held by 40 people produces a violent index. The script flags anything under its
  `--min-users` (default 50) and refuses to rank it.
- **Index without volume is a trap.** "Men's Media Fans i227" on 52 users is noise next to "Cricket Enthusiasts i161" on
  2,928. Always show both numbers on the card.
- **The population is the caveat.** The script prints the dataset's own header (site, date range, total users). Put that
  line on the slide. If the export is a media site's own traffic, the card must say the indices describe that site's
  audience, and it must not be used for an audience that does not visit it.

### 3. Build the media block from something real

In order of preference: a syndicated survey the client has (TGI, Gallup, PAS), the client's own tracker, platform data, or
— when none exist — **evidenced channel behaviour**: what the audience actually watches, comments on and shares, from the
brand's own channel data and the comment corpus. Qualitative is fine. Invented percentages are not.

### 4. Write the pain points from the corpus

Use `comment-analysis` output. For each pain point: the pain in the audience's own framing, what it means for the brand, and
the evidence (count, or a verbatim with a labelled translation). Order by what blocks the purchase, not by frequency.

### 5. Build the card

```bash
python scripts/build_tg_card.py spec.json            # writes the next free -vN
```

The spec shape is in `references/spec-example.json`, and the block-by-block geometry in `references/card-anatomy.md`.
The builder refuses to produce a card that would mislead:

- a pain point or trigger with no `evidence`;
- a number with a `%` anywhere in the media or interests block whose section has no `source`;
- an interests block whose `source` does not name a population (it looks for a dataset name and a total);
- text that would overflow its box, which is how a checked card still embarrasses you in the room.

### 6. Look at it

Render the slide to an image and read it (PowerPoint COM on Windows, LibreOffice elsewhere). Then check:

- Every number on the card appears in a source file, and every translation is labelled as one.
- The mindset paragraph would be recognised by someone from that audience, and does not read as a brand's wish.
- The pain points are things the audience said, not things the brand says about them.
- Nothing implies the interest data describes an audience it does not cover.

## Adapting to the audience

**Rural or low-digital audiences.** Analytics interest data usually does not reach them. Build identity and household from
the brief and census, pain points from the corpus, and media from evidenced behaviour. Print the gap on the card: "no
interest-affinity data covers this audience; the nearest available dataset describes urban web visitors."

**Two audiences in one brief** (a core buyer and an influencer, rural and urban). Build one card each. A single card
averaging two audiences describes nobody.

**Client-facing version.** Keep the evidence, drop the file names, and keep any caveat that changes a decision. A client
who sees "this describes our site's visitors, not farmers" trusts the rest of the deck more.

## Where TG profiles fail

- **The stock persona.** A name, a photo and adjectives ("tech-savvy, family-oriented"). It survives because nobody can
  disprove it, and it briefs nobody.
- **Borrowed data.** An interest export from a different audience, an insurance deck's media percentages on a farmer card.
- **Frequency as insight.** The biggest theme in a corpus is not automatically the biggest barrier — read the register.
- **One card for two audiences.**
- **The card that never changes.** A profile built for one campaign and reused for three years stops being evidence and
  becomes folklore. Date it, and name the data behind it, so the next team knows when it expired.

## Reference files

| File | Read it when |
| --- | --- |
| `references/card-anatomy.md` | Laying out or restyling the card; it documents each block and its geometry |
| `references/spec-example.json` | Writing a spec — copy and adapt |
| `scripts/interest_index.py` | Computing affinity indices from an interests export |
| `scripts/build_tg_card.py` | Rendering the card, with the guards above |
