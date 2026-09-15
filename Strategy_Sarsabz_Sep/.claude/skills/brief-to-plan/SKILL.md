---
name: brief-to-plan
description: Parse a client brief, RFP, tender, ToR, EOI or prequalification document and turn it into a decided plan - what the ask actually is, which requirements are pass/fail gates versus scored, what is referenced but missing from the document, and what has to start first given lead times and dependencies. Use this whenever the user shares or points at a brief, RFP, RFQ, ToR, tender, EOI, prequalification or scope-of-work document, or asks "what are they actually asking for", "what do we need to submit", "where do we start", "what should we do first", "can we bid on this", or is opening any new pitch or submission - even if they only say "have a look at this brief".
---

# Brief to plan

Turn a brief into two things: **a statement of what is actually being asked**, and **a sequenced
plan of what to do first**.

Briefs, and tender documents especially, bury the ask. A 34-page procurement document is
typically 90% boilerplate — procurement law, fraud clauses, grievance mechanisms — wrapped
around two pages that actually decide whether you win. The work here is separating those two
pages from the rest, then working out what has to start today.

## The two failure modes

**Reading only what is present.** The most consequential thing in a brief is often what it
references but does not contain — a schedule of requirements listed in the contents and absent
from the file, an annexe that lives on the portal, a brief the buyer assumes you have. Step 2
exists entirely for this.

**Sequencing by marks.** The instinct is to start with the highest-scoring item. That is wrong
when a low-mark item has a long external lead time. A bank certificate may be worth two marks
and take a week of somebody else's time; started on day six of an eight-day window, it fails,
and a failed gate makes every other mark worthless. **Sequence by latest possible start, not by
value.**

## Workflow

### 1. Extract the document

```bash
python scripts/extract_brief.py <brief.pdf> --out <workspace>/brief
```

Handles PDF and DOCX. Writes the full text, a section-by-section split, a report of internal
cross-references that have no matching section, and any dates it can find.

Read `references/extraction-gotchas.md` if the output looks thin or wrong — portal-generated
PDFs in particular defeat naive extraction, and a silently half-extracted brief is worse than
no extraction.

### 2. Find what is missing before reading what is there

Check the script's `missing_sections` report, then verify by hand:

- Does the table of contents list sections the body does not contain?
- Do clauses cross-reference annexes, schedules or lots that are not in the file?
- Does the document point at a portal, a data room, or "as advertised"?
- Are there addenda or clarification responses issued after the version you hold?

Every gap becomes a **question for the buyer**, and questions have a deadline of their own —
usually days before submission and easy to miss. Getting the clarification window into the plan
is often worth more than any analysis you could do instead.

Do not fill a gap with an assumption and proceed quietly. Name it, ask it, and plan for both
answers.

### 3. State the ask in one paragraph

Before any register or spreadsheet, write plainly: **who wants what, for whom, by when, and how
will they choose.** If that paragraph is hard to write, the brief is ambiguous — which is itself
the finding, and the basis of your clarification questions.

Include what the buyer is *really* buying, which is not always what the document's title says. A
document headed "prequalification" that allocates 70 of 100 marks to a speculative campaign is
not a paperwork exercise; it is a pitch with a paperwork gate.

### 4. Classify every requirement

Three types, and the distinction drives everything after:

| Type | Meaning | Consequence |
|---|---|---|
| **gate** | Pass/fail eligibility | Failing one voids the entire submission regardless of score |
| **scored** | Carries marks | Effort should follow marks — but only after gates are safe |
| **informational** | Context, terms, boilerplate | Note and move on |

For each, capture the fields in `references/register-schema.md`. Two matter more than the rest:

- **`external`** — is this in someone else's hands? Certificates, attestations, bank letters,
  client references, notarised declarations. External items dominate the schedule because you
  cannot compress them by working harder.
- **`depends_on`** — what must finish before this can start. This is where the real sequence
  lives. Category and competitor analysis typically feed concept, strategy and channel
  recommendations, so it starts early despite carrying fewer marks itself.

Write these into `requirements.yaml`.

### 5. Generate the plan

```bash
python scripts/plan_from_register.py requirements.yaml --out <workspace>/brief
```

Back-plans from the deadline through the dependency graph and produces:

- **Start today** — items whose latest start has arrived or passed
- **Critical path** — the chain with no slack
- **Overdue** — items that were already too late when you started, flagged rather than hidden
- **Marks per day** — where effort buys the most score, for allocating people
- A register spreadsheet and a dated schedule

An overdue item is not a reason to despair; it is a reason to escalate, compress or ask for an
extension *now* rather than discover it on the last day.

### 6. Decide, and say what you decided

The plan is an input to a judgment, not the judgment. Close with:

- **Bid / no-bid.** If a gate cannot be met — an accreditation you do not hold, a threshold you
  do not clear — say so immediately and plainly. That conversation is cheap on day one and
  expensive on day seven.
- **What starts today, and who owns it.** Named people, not functions.
- **What we are asking the buyer**, and by when.
- **What we are deliberately not doing.** A plan that fits the time available requires
  something to be dropped; make that a decision rather than an accident.

## Handing off

Once the plan exists, the analysis work usually starts with the category. The
**`competitor-comms-audit`** skill covers that end: competitive set, ad audit, message-territory
map and the vacant-territory finding that a campaign concept is built on. Trigger it as soon as
the plan says the category work starts, and pass it the ask statement from step 3 so the audit
is scoped to the actual brief rather than the category in general.

## Output structure

```
1. The ask            - one paragraph: who, what, for whom, by when, judged how
2. How they choose    - gates, scored criteria with weights, where the marks concentrate
3. What is missing    - referenced-but-absent, and the questions we are filing
4. The sequence       - start today, critical path, dated schedule
5. The call           - bid/no-bid, owners, what we are not doing
```

Sections 3 and 5 are the ones that get skipped and the ones that decide outcomes. Do not drop
them because the deadline is tight — a tight deadline is precisely when an unasked question or
an unmet gate costs the most.
