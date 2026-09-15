# Writing the analysis, section by section

The structure in SKILL.md is a spine, not a form to fill. Sections earn their place by
having evidence behind them; a section with nothing in it should say so in one line and
get out of the way rather than be padded.

## 1. Executive summary

Four or five sentences, each of which a reader could act on. Lead with the thing that
changes a decision, not with the size of the corpus.

Weak: "Consumers discussed a range of topics including speed, price and service."
Strong: "The corpus is dominated by people who already have the product and cannot make
it work. Speed is mentioned in 158 comments, but the complaint underneath it is almost
always resolution, not bandwidth."

Name the single biggest barrier and the single biggest asset, and say what the brand
should do about each.

## 2. What consumers are trying to solve

The job, in their words. Bullets, each a real motivation visible in the corpus — not
personas, not demographics. If the corpus shows people shopping, this is a shortlist
question. If it shows owners, it is a resolution question. Those are different briefs.

## 3. Pain points  [table: Pain | Consumer implication]

One row per distinct pain, ordered by weight of evidence. The second column is the
*implication*, not a restatement: what it means for the brand that this is what people
say. A row that merely rephrases the pain is wasting space.

Put the count in the row where you have it, and a verbatim under the table for the
heaviest two or three.

## 4. What consumers seek  [table: What they seek | What it means for the brand]

The positive mirror of section 3. Seek ≠ absence of pain: people ask for things they
have never had. This is where the campaign territory usually comes from.

## 5. What they dislike

Short, sharp bullets. The avoid-list for creative — the things that will actively cost
the brand credibility if it says them. Often the most useful page for a creative team.

## 6. Relative importance of themes  [table: Theme | Comments | Share]

Straight from `theme_counts.csv`. Two things must appear with it:

- **Themes are not mutually exclusive.** One comment can carry several; the column does
  not total to 100%.
- **Coverage.** What share of the corpus matched any theme at all.

Then interpret: which themes are unusually prominent for this category, and which are
conspicuously absent. Absence is often the finding — a category that never discusses
price is telling you price is not the battleground.

## 7. Brand associations  [table: Brand | Mentions | Explicit + | Explicit −]

From `brand_sentiment.csv`. Carry the caveat in the text immediately under the table, not
in a footnote: this is rule-based coding of explicit language, it is directional rather
than a sentiment measurement, and most mentions carry no explicit marker either way.

Never rank brands by a computed sentiment percentage as though it were measured.

## 8. What each brand owns  [table: Brand | Positive association | Negative / barrier]

This is a qualitative read, written after reading the comments for each brand — not
generated from counts. Two or three specific associations per brand, in the commenters'
own language where possible. A brand with too few mentions to characterise gets a row
saying exactly that.

## 9. Client deep dive

Read *every* comment mentioning the client. State the count and the share of the corpus.
Then the qualitative read, with verbatims — the advocacy, the objections, the specific
numbers people cite, and the gap between claim and reported experience.

If the client is barely mentioned, that is itself the headline, and it should be said
plainly rather than dressed up.

## 10. SWOT

Only where the evidence reaches. See the SWOT section in SKILL.md — the governing test is
whether you could show the client the comments each bullet came from. Label inferences as
inferences. A thin box stated honestly beats four invented bullets.

## 11. Strategic implications  [table: Direction | Why it matters]

Decisions, not observations. "Sell confidence, not features" is a direction; "consumers
care about confidence" is a restatement of section 3. Each row's second column ties back
to specific evidence.

## 12. Questions for the strategy team

The things the corpus raises but cannot answer, and which only the client can. This
section is what separates an analyst from a summariser: it converts every gap you refused
to invent into a productive next step.

## 13. Data & methodology

Always include. State:

- Source file, and the date collected or supplied.
- **Comments analysed**, and how that differs from rows in the file (scaffolding removed).
- Reply share, script/language mix.
- How coding was done, how many themes, and the coverage share.
- The sentiment method and its limits, in one sentence.
- The limits, which must always include: the corpus is self-selected and is not a sample
  of the customer base; it cannot support satisfaction rates or market share; and
  comments report what someone wrote, not verified fact.
