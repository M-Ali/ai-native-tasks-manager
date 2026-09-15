---
name: comment-analysis
description: "Turn a file of YouTube (or other social) comments into a decided consumer analysis - the consumer pulse, pain points, what people are seeking, what they dislike, how each brand is talked about, and a SWOT of the brand under discussion. Use whenever the user shares or points at a file of comments, reviews, or social conversation in .xlsx, .csv or .txt and wants to know what consumers are saying, what the pain points are, what the consumer pulse or sentiment is, what people want, how a brand is perceived, or asks for a SWOT, voice-of-customer, consumer insight or conversation analysis - even if they only say 'here are some comments, what do you make of them' or 'can we get anything out of this'. Also use when a category scan, pitch or brand review needs demand-side evidence to sit alongside what brands are publishing."
---

# Comment analysis

A comment dump is the cheapest primary consumer research a brand will ever get: unsolicited,
unmoderated, already written in the customer's own words. It is also usually skimmed and
quoted selectively, which is how a deck ends up asserting a mood rather than evidencing one.

This skill turns that file into an analysis with numbers behind every claim, verbatims next
to every finding, and an explicit account of what the corpus cannot tell you.

## The one thing to get right

**Frequency is how you find the finding. It is never the finding itself.**

A count tells you what your lexicon caught. It cannot tell you what people mean, and the
two often point in opposite directions. In one real corpus, `speed` was the largest theme
at 158 comments — but reading those comments showed almost none were about bandwidth.
They were about paying for 20MB and receiving 5, which is a complaint about *honesty*
wearing a technical costume. An analysis led by the frequency table would have
recommended a speed campaign to a brand whose actual problem was that customers believed
it was cheating them.

So the counts belong in the method section and in the ranking of where to *look*. The
analysis itself is built from what people actually say. Concretely:

- **Lead every section with the substance, not the number.** "Commenters describe a bill
  that changes without warning" — then the evidence. Not "85 comments mention price."
- **A single articulate comment can outweigh a hundred vague ones.** Someone explaining
  *why* they left, with amounts and dates, is worth more than fifty people typing "bad
  service". Weight by explanatory power, not by volume.
- **Read the emotional register, not just the topic.** "Slow internet" and "ye log fraud
  hain" can both code as a complaint. One is a performance issue; the other is an
  accusation of dishonesty, and no amount of service improvement answers it. The register
  tells you what kind of problem the brand has.
- **Notice what is said about the brand versus what is said about the product.** "The
  connection is fine in my area" is not brand advocacy — it is a statement about a cable.
  A brand whose defenders only ever defend their individual line has no equity
  independent of local network quality, and that is a finding a frequency table cannot
  produce.
- **Absence of a sentiment is evidence too.** If nobody in a thousand comments says a
  brand is *fair*, or *modern*, or *on my side*, that silence is louder than any count.

The test before a section ships: *does this tell the client something a word-frequency
chart could not?* If not, it is not yet an analysis.

**Say what the evidence does not cover.** A comment corpus is self-selected: people who are
angry, people who are shopping, and people being nice to a creator. It is not a sample of
the market, and it cannot support a satisfaction rate, a market share, or a claim about
"most customers".

The temptation is to fill gaps — invent a demographic split, estimate sentiment as a
percentage, name a cause for a complaint. Don't. **A gap stated plainly is a finding; a gap
filled by assumption is a liability**, because the one number a client checks is the one
that is wrong, and it discredits the rest of a sound analysis.

When something is genuinely not in the data, write one of these and move on:

- "The corpus does not indicate X."
- "N comments touch this — too few to generalise, quoted here as signal not proof."
- "This is a question for the client, not a number to estimate."

## Workflow

### 1. Load the file and find out what is actually in it

```bash
python scripts/load_comments.py <file.xlsx|.csv|.txt> --out <workspace>
```

Pasted exports are not tabular. Two layouts turn up, often in the same file: a **block**
layout where each comment is followed by its like count, author and timestamp as separate
*rows*, and a **flat** layout of one comment per row, with replies in a second column.
Treating every non-empty cell as a comment inflates the corpus several-fold and fills the
analysis with rows reading "Reply", "1" and "@username".

The loader strips that scaffolding, keeps replies (flagged, because a reply is usually a
rebuttal or corroboration — exactly what you want when reading a debate), and writes
`load_report.md` with the arithmetic: cells in, scaffolding out, comments analysed.

**Quote "comments analysed", never the row count.** In one real file, 1,099 rows and 1,068
non-empty cells produced **976 actual comments**. Reporting 1,099 would have overstated the
base by 13% in a document whose credibility rests on its numbers.

### 2. Build the lexicon for this category

```bash
cp assets/lexicon.example.yaml <workspace>/lexicon.yaml   # then edit it
```

This is the highest-leverage twenty minutes in the whole job, because **an unmatched theme
is invisible in the output**. Read fifty comments first and put the category's own words in
— including the way people actually type them.

Three things that decide whether this works:

- **Single-quote every regex.** YAML processes escapes inside double quotes, so
  `"\bptcl\b"` arrives as two *backspace characters* and matches nothing, silently. That
  exact mistake reported 11 mentions of a client brand in a corpus holding 92 — it would
  have concluded the brand was barely discussed when it was the most discussed brand in the
  file. `code_comments.py` now aborts if it sees a control character in a pattern, but
  write `'\bptcl\b'` and the question never arises.
- **Include the code-switched spellings.** In markets that mix scripts, `ghatiya`,
  `bakwas`, `masla`, `achi` and the native-script equivalents carry a large share of the
  strongest opinion. Omit them and you silently drop the angriest part of the corpus.
- **Code `video_feedback` separately.** A large slice of any YouTube corpus is praise for
  the *video* — "great video", "informative", "thanks bro". That is audience feedback to a
  creator, not consumer voice about the category. Coding it lets you exclude it from pain
  points and say what share of the file it was.

### 3. Count what is there

```bash
python scripts/code_comments.py <workspace>/comments.csv --lexicon <workspace>/lexicon.yaml --out <workspace>
```

Writes `coded.csv`, `theme_counts.csv`, `brand_counts.csv`, `brand_sentiment.csv`,
`cooccurrence.csv` and `coding_report.md`.

The coding is deterministic rather than model-read, for one reason: **the counts get
quoted**. "54 comments raise service" ends up on a slide, so it has to be reproducible on a
rerun and traceable to the comments that matched. A model reading a thousand comments
resamples and gives a different number each time, which cannot be defended in a room.

**Check coverage before you use any count.** The report states what share of comments
matched at least one theme. Typically 35-60% for YouTube — the remainder are short
affirmations, greetings and off-topic chatter with no codeable subject. If it is under
~35%, read forty unmatched comments and extend the lexicon; the problem is almost always
the vocabulary, not the corpus.

**Sentiment here is directional, not measured.** The script counts explicit positive and
negative markers near a brand mention, including negated positives ("not good", "sahi
nahi"). Sarcasm and code-switching defeat it, and most mentions carry no marker at all. Use
it to say *"Oshan is talked about warmly, MG is not"*. Never publish "MG has 10.7% negative
sentiment" as a measurement.

### 4. Read the comments — this is where the analysis actually happens

Everything before this step was sorting the post. The scripts tell you how often a subject
appears and how it skews; they cannot tell you what it means, which comment is the telling
one, or what the brand should do.

Read every comment mentioning the client — all of them, however many — and read the
heaviest two or three themes in full. Use the counts only to decide reading order.

As you read, you are looking for **what people are saying about the brand**, which is a
different question from what they are saying about. Ask of the corpus:

- **What kind of relationship do they describe having with this brand?** Chosen, tolerated,
  endured, escaped? A customer who says "I had no other option" is telling you the brand's
  market position in six words.
- **What do its defenders actually defend?** If they only defend their own connection, the
  brand has no equity of its own.
- **What register do the complaints use?** Functional ("it's slow"), moral ("they're
  thieves"), or resigned ("there's no point complaining")? These are three different
  briefs.
- **What are they saying to each other rather than to the brand?** A comment section where
  buyers ask strangers for advice is one where brand claims have stopped being believed.
- **What does nobody say?** The missing compliment is often the positioning gap.

Pull verbatims as you go — **every finding needs a real quote sitting next to it**, because
the quote is what makes a marketing director believe the point.

Look specifically for:

- **The unprompted comparison.** "I watched this and still chose X" is worth more than fifty
  generic compliments: it is a shopper narrating their own decision.
- **The person explaining why they left.** A departure narrated with dates and amounts is
  the most valuable single artefact in any corpus.
- **The specific number a customer states.** "550 km per tank", "20 Mbps ka package aur 2
  Mbps aata hai", "4 saal se sahi hai". Specifics are checkable and they anchor a section.
- **The gap between what a brand claims and what an owner reports.** That gap is usually the
  single most valuable observation in the file.
- **Questions.** A corpus full of "is it available in my area?" is a distribution problem
  wearing a communication costume; a corpus full of "how do I fix this?" is a service
  problem. They lead to different recommendations.

### 5. Write it as an argument

Write the analysis as Markdown, then render it. Markdown stays the source so the working
file and the delivered file cannot drift:

```bash
python scripts/render_docx.py <workspace>/analysis.md --out <workspace>/out
```

Use this structure. It moves from what people want, through what is wrong, to what the
brand should do — and it puts method and limits at the end where they belong, stated rather
than buried:

```
# <CLIENT> — Consumer Conversation Analysis
Analysis base: N comments analysed (from M rows)

1.  Executive summary            the four or five things that matter, up front
2.  What consumers are trying to solve    the job, in their terms
3.  Consumer pain points          [table: pain | consumer implication]
4.  What consumers seek           [table: what they seek | what it means for the brand]
5.  What they dislike             the avoid-list, sharply
6.  Relative importance of themes [table: theme | comments | share] + interpretation
7.  Brand associations            [table: brand | mentions | explicit +/- ] + caveat
8.  What each brand owns          [table: brand | positive | negative/barrier]
9.  <CLIENT> deep dive            every client mention read, with verbatims
10. SWOT                          only if the evidence supports it — see below
11. Strategic implications        [table: direction | why it matters]
12. Questions for the strategy team   the things only the client can answer
13. Data & methodology            base, coding, coverage, and the limits
```

Adapt it. If the corpus has no competitor mentions, sections 7 and 8 collapse into one
paragraph saying so. If it is all service complaints and no purchase consideration, say that
— a narrow corpus honestly described is more useful than a wide one padded out.

### 6. SWOT — only where the evidence reaches

A SWOT is where invention creeps in, because the four boxes invite filling. Hold the line:

- **Strengths and Weaknesses** must come from the corpus — what people actually praise and
  complain about, with counts and verbatims behind them.
- **Opportunities and Threats** may draw on competitor mentions *in the corpus*, plus
  anything the user has supplied. Anything else is your inference and must be labelled as
  such.
- **If a box is thin, say so.** "The corpus supports two weaknesses; a third is likely but
  not evidenced here" is a stronger sentence than four invented bullets.
- **If the brand is barely mentioned, do not build a SWOT at all.** Write instead: "N
  comments mention <brand>, too few for a defensible SWOT. What they do show is ___." Then
  say what would be needed.

The test for every bullet: *could I show the client the comments this came from?* If not,
cut it or label it an inference.

## Things worth checking that clients rarely have

- **Reply-to-comment ratio.** A corpus that is a third replies is a debate, not a comment
  section, and the argument being had is usually the category's real purchase question.
- **Language and script mix.** `load_report.md` counts it. A brand advertising only in
  English to a corpus that argues in Urdu has a relevance problem it has not noticed.
- **Who gets defended.** Unprompted defence of a brand by a stranger ("mera toh 4 saal se
  sahi hai") is advocacy, and it is rarer and more valuable than praise.
- **Which complaints repeat verbatim.** The same phrasing recurring across many commenters
  is a systemic failure, not a set of unlucky individuals.

## Rules that keep this honest

- **Never invent a number.** If it is not in the corpus or supplied by the user, it does not
  go in the document. Say the corpus does not cover it.
- **Quote the base correctly** — comments analysed, not rows in the file.
- **Every claim gets a count or a verbatim**, ideally both.
- **State coverage** — what share the taxonomy actually matched.
- **Sentiment is directional.** Never present it as a measurement.
- **Separate audience feedback from consumer voice.** Praise for the video is not praise for
  the brand.
- **A comment is what someone wrote, not a fact.** Report "commenters report X", never "X is
  the case" — particularly for claims about faults, billing or safety.
- **Small numbers stay small.** Under ~15 mentions, quote it as signal and say it is thin
  rather than converting it to a percentage.

## Reference files

| File | Read it when |
| --- | --- |
| `assets/lexicon.example.yaml` | Starting any run — copy and adapt it |
| `references/output-template.md` | Writing the document; full section-by-section guidance |
| `scripts/load_comments.py` | First step — parses xlsx/csv/txt, strips scaffolding |
| `scripts/code_comments.py` | Second step — deterministic counts and directional sentiment |
| `scripts/render_docx.py` | Last step — Markdown (with tables) to .docx |
