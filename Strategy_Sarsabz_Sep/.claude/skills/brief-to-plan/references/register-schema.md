# requirements.yaml schema

One file per brief. The model fills it by reading the brief; `plan_from_register.py` turns it
into a dated plan. Judgment goes here; arithmetic goes in the script.

## Shape

```yaml
brief:
  client: State Life Insurance Corporation of Pakistan (SLIC)
  reference: P78118
  title: Prequalification for Procurement of Advertising Agencies
  source_file: data/pq-bidding-document (2).pdf
  portal: https://epads.gov.pk/opportunities/federal/procurements/78118
  method: Single Stage - One Envelope, National
  selection: Quality Based Selection (QBS)
  total_marks: 100
  pass_mark: 50
  submission_deadline: 2026-09-03 11:00
  submission_channel: EPADS v2.0 only - manual submission not entertained
  clarification_deadline: 2026-08-31

the_ask: |
  One paragraph, plain language. Who wants what, for whom, by when, and how will
  they choose. Say what is really being bought, which is not always the title.

missing:
  - what: Section VII Schedule of Requirements
    referenced_at: ITA 6.8; PDS clause 1 points to "section services and Lots"
    action: Check the portal, then file a clarification
    question: Which product line, audience and objective should the campaign address?

requirements:
  - id: apns_pba_pid
    name: APNS, PBA and AAP registration plus PID enlistment
    type: gate            # gate | scored | informational
    marks: 3              # 0 for pure gates
    effort_days: 1        # our working days
    lead_time_days: 5     # time in someone else's hands
    external: true        # do we control the outcome?
    owner: Compliance
    depends_on: []        # ids that must finish first
    notes: Not evidenced anywhere in prior submissions. Confirm before anything else.
```

## Fields

| Field | Why it matters |
|---|---|
| `id` | Referenced by `depends_on`. Short, stable, lowercase. |
| `type` | `gate` items are pass/fail — the plan flags them regardless of marks. |
| `marks` | Drives the marks-per-day ranking. Zero is valid for a pure gate. |
| `effort_days` | **Our** working days. Compressible by adding people, sometimes. |
| `lead_time_days` | **Someone else's** time — a bank, a regulator, a notary, a client reference. Not compressible. This is what makes low-value items urgent. |
| `external` | Anything you cannot finish by working harder. External items get surfaced first in the plan. |
| `owner` | A named person, not a department. Unowned work does not happen. |
| `depends_on` | The dependency graph. Where sequencing actually comes from. |
| `due` | An earlier deadline of the item's own — a clarification window, a site visit, a registration cut-off. Without it the item back-plans to the submission date and looks less urgent than it is. |
| `notes` | Anything the schema cannot hold — especially *why* a status is what it is. |

## Estimating the two durations

Keep them separate; they behave differently.

**`effort_days`** is working time. Be honest rather than optimistic — a plan built on
best-case effort produces a schedule that is wrong from day one, and everyone stops trusting
it by day three.

**`lead_time_days`** is waiting time. Ask the person who has actually requested that document
before, not the official turnaround. A tax certificate might be quoted as three days and take
ten in practice.

Where an item is pure waiting — you submit a form and wait — put the wait in `lead_time_days`
and a small `effort_days` for the submitting.

## Dependencies worth modelling

The ones that recur across pitches:

- Category and competitor analysis feeds **concept**, **communication strategy** and **channel
  or media recommendations**. It carries fewer marks than the things it feeds, and therefore
  looks skippable, and therefore gets started too late. Model it explicitly.
- Concept feeds **creative execution**. Artwork started before the concept lands gets thrown
  away.
- Team and CV material feeds both a **staffing annexe** and the **capability narrative**. Gather
  once.
- Everything feeds **assembly**, which is always underestimated. Give it real days.

## Common mistakes in filling this in

**Marking something internal because your colleague controls it.** If you cannot finish it
alone, the schedule risk is real. `external: true`.

**Omitting assembly, proofing and upload.** Portal uploads fail, files exceed size limits,
signatures need wet ink. Model submission as a requirement with its own effort and a buffer.

**Modelling the deadline as the finish.** Finish before it. `plan_from_register.py` applies a
buffer by default; do not set it to zero because the schedule looks tight — a schedule that only
works if nothing goes wrong is not a schedule.
