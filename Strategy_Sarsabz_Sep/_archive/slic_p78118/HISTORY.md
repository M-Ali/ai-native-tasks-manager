# HISTORY

Newest first.

## 2026-08-28 (housekeeping) — `slic-pq doc`; duplication removed; docs squared

**Removed 300 lines of duplication.** `build_concept_doc.py` and `build_strategy_doc.py`
were 150-line copies differing in five naming lines — and the same two rendering bugs had
already been fixed twice, once in each. Extracted to `src/slic_pq/submission_doc.py` and
exposed as **`slic-pq doc <markdown> --stem <name>`**. Verified byte-parity (3 and 6 tables,
84 and 87 paragraphs, zero stray asterisks) before deleting both builders.

The renderer documents the two rules that were bugs first: list items and blockquotes wrap
across source lines and must be joined before rendering, and nested emphasis is unsupported
because it is ambiguous in Markdown itself.

**README** gained a campaign-submission map (which criterion, which source file, what state
it is in), the `slic-pq doc` workflow, the two rendering rules, and `submission_doc.py` in
the layout.

**AGENTS.md** gained six conventions for writing scored sections — Markdown is the source
and Word a render, verify figures before writing not after, state assumptions on the page,
say what the evidence cannot show, distinguish a data request from an assumption, and check
heritage claims before artwork.

## 2026-08-28 (strategy) — `communication-strategy` skill; SLIC strategy drafted

**New skill.** Five parts, matching what tenders actually ask for: communication objective,
target audience, strategic approach, message architecture, communication journey. Built
around three disciplines the category gets wrong — a business objective is not a
communication objective (write it from -> to, and name which of the five barriers it
clears); an audience defined by demographics is a media buying instruction, so separate the
primary audience from influencers and the **decision-making unit**; and a proof layer full
of adjectives proves nothing. Journeys are **barrier-led**, not generic funnels: a standard
awareness -> consideration model assumes the barrier is awareness.

**SLIC strategy drafted** — `workspace/strategy/the_strategy.md` +
`SLIC_P78118_communication_strategy_2026-08-28-v3.docx`.

The strategic call worth recording: **category-building, not recruitment.** Penetration is
3% of the population and State Life holds **55% share** (CCP 2025), so the majority of any
category growth accrues to State Life. Fighting EFU or Jubilee for share means fighting over
3% of the country. The competitor is non-consumption.

Barrier diagnosed as **belief, not awareness** — 233 of 288 ads prove nothing, 14 address
the objection. Consequence stated plainly: more reach will not fix this, so the money belongs
behind proof rather than frequency.

**An honesty section opens the document:** the audit is a supply-side read of what the
category *says*, not consumer research. The barrier diagnosis is an inference, labelled as
one, with a claims-attitude tracking wave recommended rather than assumed.

Also recorded: pillar 1 rests on **complaints resolved**, which is what State Life publishes
today — not claims paid. If a claims-paid figure exists it outranks Jubilee's PKR 54 billion.
Raised as a **data request, not an assumption**.

Two more markdown-to-Word rendering bugs fixed in both doc builders: emphasis spanning a
wrapped blockquote line, and nested emphasis (un-nested at source, since it is ambiguous
markdown anyway).

Five skills now.

## 2026-08-28 (concept) — Campaign Concept drafted, four ranked directions

`workspace/concept/the_concept.md` + `SLIC_P78118_campaign_concept_2026-08-28-v3.docx`.

**Corrected a fact before it reached artwork.** State Life is **54 years old**, not 60 —
established under the Life Insurance (Nationalisation) Order of 19 March 1972, consolidated
management from 1 November 1972. Its own creative says "Since 1972" and "For over five
decades". A longer claim is available but must be worded as inherited: State Life was formed
by merging **32 companies**, several predating Partition. "60 years of State Life" is
checkable and false.

**Four directions, ranked against the skill's six tests**, each named with its proper
execution format rather than a vague description:

1. **The Promise You Can Check** — reason-why, supported by documentary testimonial.
2. **Aye Khuda Meray Abbu Salamat Rahain** (revival) — emotional drama, slice-of-life.
3. **Trusted by Your Parents** — two-generation vignettes. Ranks third because heritage is
   the second-most-crowded territory: 29 of 288 ads claim national trust, but 22 of those
   sit in the pride register.
4. **Claim What Is Yours** — problem-solution as direct response. A segment burst, not a
   platform; cannot carry the other 55 marks.

**Recommendation: A as the platform, B as its line.** A has the argument and no warmth;
B has the warmth and no argument. The line is a prayer; the platform's job is to answer it.
Mechanic: the E-Kachehri resolution figure, published every cycle — it refreshes itself, so
the platform generates work rather than repeating a launch.

Every figure re-verified against `capture.xlsx` before writing: 233 no-proof, 3
claim_statistics, 14 objection-handling, 2 boomer_plus, 29 national_trust of which 22 pride.

The missing Section VII assumption is stated on the page — concept built at corporate brand
level so it can carry any product line, rather than a product concept that cannot widen.

`build_concept_doc.py` renders the markdown so the two cannot drift. Two rendering bugs
fixed: emphasis markers leaking into headings, and emphasis spanning a wrapped line inside a
list item losing half its markers.

## 2026-08-28 (later still) — scan skill upgraded; `campaign-concept` skill added

**Scan skill upgraded** from what this project taught it. `analyze_scan.py` gained three
standard reads: **9a/9b script actually used in captions**, measured on the raw text rather
than the `language_lead` code — because a coded field with a default measures the coder, not
the category, and a naive tally would have reported 97% English when that was simply the
default; **4b output that sells nothing** (calendar share per brand); and **9c investment vs
who appears in the work** (studio shoots against customers and executives in frame).

SKILL.md gained a step 9, **Find the why, not just the what** — a pattern is an observation,
the mechanism is the insight, and it is usually documented publicly rather than inferable.
Two tests: can you cite it, and does it change the recommendation. Plus a table of the eight
optional `reads.yaml` slide keys the deck builder now renders, and a new honesty rule about
never counting your own defaults.

**New skill: `campaign-concept`.** Converts an audit finding into an idea. Built around one
chain — asset → tension → territory → proposition → line → executions — and two tests: could
a competitor run this line tomorrow, and does the concept still stand if you delete the
audit. `references/concept-tests.md` carries six kill-tests (swap, asset, deletion,
department, month-twelve, room) and a table mapping a concept to separately-scored tender
criteria, since a concept that cannot feed artworks, strategy and media is a slogan.

Two lessons from this bid written in: prefer reviving a line the client already owns over
inventing one, and where a brief omits product or audience, state the assumption on the page
rather than leaving it implicit.

Four skills now: `brief-to-plan`, `category-creative-scan`, `competitor-comms-audit`,
`campaign-concept`.

## 2026-08-28 (later) — the "why" behind app convergence; target group; market share

The scan said every brand is converging on an app. It did not say **why**, which is the
difference between an observation and an insight. The mechanism turns out to be documented,
not inferable: **none of these insurers owns its customer.**

From the **Competition Commission of Pakistan, State of Competition Report 2025**: banks
"impose additional internal limits on the amount of business insurance companies can conduct
through banks" — which the CCP calls a **refusal to deal**; and bank/insurer staff "do not
properly guide the customers", with terms in fine print. Add the market context — **3% of
Pakistanis hold a life policy**, penetration **0.67% of GDP** vs 6.7% globally, density
**US$14** vs India's US$82 — and the app is rational: a direct relationship the bank cannot
gate, and a reason to open something between annual premium notices.

**It solves the insurer's distribution problem, not the customer's trust problem.**

The strategic payoff: **State Life holds 55% share (CCP) and distributes through a tied
agency force, not banks.** It already owns the relationship the others are buying an app to
get. "Don't chase the app" is now a structural argument rather than taste.

That report also supplied the **market-share figure the SOV-vs-SOM chart had been missing
since the start** — now cited in `capture.xlsx` MarketShare and in the findings.

**Target group added.** Millennial 35% · multi-generational 31% · Gen X 23% · Gen Z 10% ·
**60+ 1%**. The audience follows the distribution: the work is cast for the banked, urban,
app-owning customer bancassurance already reaches. Two ads in 288 address the over-60s, both
State Life's.

**Deliverables:** deck now 20 slides with a "Why they are all building an app" slide, a
target-group slide, and a red WHY line on every brand slide. `build_deck.py` gained `_why`,
`_audience` and per-brand `why` support. `the_finding.md` gained §2a and §2b; the Word
diagnostic regenerated to v2 (9 tables) and pulls both new sections from the finding, so the
three documents cannot diverge.

## 2026-08-28 — Word diagnostic; README brought current

**Added the missing deliverable.** The scan had a deck but no document. SLIC scores every
criterion `(Qualitative)(Doc Required)` — each scored line needs an upload, and a slide
pack alone is a thin answer to 15 marks. `workspace/audit/build_diagnostic_doc.py` builds
`SLIC_P78118_competitive_analysis_<date>.docx`: the finding in prose, then seven tables an
evaluator can recount, then a method-and-limits section that states what the evidence
cannot show.

The prose is parsed out of `the_finding.md` rather than retyped, so the document and the
finding cannot drift apart. Tables are rebuilt natively from `capture.xlsx`.

**README.md was two days stale** — it documented only the `slic-pq` CLI and none of the
scan. Now carries the full seven-step scan workflow with runnable commands, the skills
table and sync instruction, and the current `workspace/` layout.

Three skills to date: `brief-to-plan`, `category-creative-scan`, `competitor-comms-audit`.

## 2026-08-27 (later) — deep read: 288 posts, finding Draft 2, deck v4

**The 12-post window was wrong about the client.** Instagram's logged-out ceiling is 24,
not 12 — the grid loads 12, a "Show more posts" button yields one more tranche, then it
disappears (traced round by round). Doubling the read overturned the central judgement.

Draft 1 said State Life's channel "reads as a ministry newsletter" with its proof "filed as
paperwork". At 24 posts that is false. State Life runs **Policyholder's Voices** (named
claimants Sara Rizvi and Mahin Khan), **E-Kachehri published twice** with the resolution
figure on the creative (745/744, then 730/729), a **"Since 1972"** heritage claim, and the
revived fifty-year line **"Aye Khuda Meray Abbu Salamat Rahain"**. All of it sat just
outside a fortnight's window.

The correction strengthened the case: from *"State Life should start proving"* to *"State
Life already proves, better than anyone in the category, and then hides it."* Two of the
three performance numbers published across 288 ads are State Life's own; so are the only
two ads cast for the over-60s.

Numbers on 288: **81% carry no proof device**, 3 cite a performance figure, 14 (5%) address
the core objection, 4 brands publish **zero** brand-building, Jubilee drives 79% of output
into its app. State Life is 38% brand-building, second to Askari's 71%, and the only brand
running two campaign platforms at once.

**Skill hardening from real use.** Ten bugs found by the eval runs and fixed, three of them
data-corrupting: `reposted_from` returned "p"/"reel" for ordinary posts; pinned posts made
an active brand read as 0.6 posts/week; a colon in `--title` silently wrote a 0-byte deck
into an NTFS alternate data stream. Added `seed_capture.py` and `apply_codes.py` (both
agents had hand-rolled that glue), `select_for_slides.py`, `--depth`, `--enrich`,
`--claims`, `--append`, a DORMANT cadence band, and a burst flag. Post dates and full
captions turn out to be readable from the post page without a login; likes are not (tested
three posts, all null).

**Slide selection is now a stated rule**, not taste: every proof-device ad, every
campaign-platform ad, then stratified to the brand's content-type mix. `slide_selection.csv`
records the reason per tile.

Caveats carried into the deliverables: share of voice still unmeasurable (equal N), nothing
paid, 79 of 288 posts uncaptioned and coded to a brand-typical default, `@pqftl_official`
unverified at 129 followers.

## 2026-08-27 — `category-creative-scan` skill; 144-ad scan; brand-wise deck

**Ran the category scan.** 144 posts across 12 brands, every image downloaded and coded
from the artwork. `workspace/audit/` holds `capture.xlsx`, `media/`, `contact/`,
`the_finding.md`; deck at `workspace/out/SLIC_P78118_category_audit_2026-08-27.pptx`.

Access findings, all tested rather than assumed: Facebook pages return the title only;
`instaloader` anonymous gets **HTTP 429 on the first call**; the `/embed/captioned/`
endpoint is now login-walled; `curl` gets a 615 KB shell. **A headless Chromium
(`browser-automation`'s `browser.mjs`) renders Instagram fine** and yields followers, bio,
verified flag and ~12 posts with downloadable 640px CDN images — no login, cookie, card or
extension. Facebook post *text* is recoverable through web search even when the page will
not load; that recovered a verbatim competitor caption with a citable URL.

The finding: **78% of category ads use no proof device**, two of 144 cite a claims number,
and 6% address the core objection. State Life already owns the strongest credibility number
in the category — *744 of 745 complaints resolved* — and publishes it as grievance-redress
paperwork.

**A correction worth recording.** An early read said State Life's feed had "no product, no
proposition" — based on four of twelve images. The full grid contains a real campaign
(*PLAN AHEAD WITH STATE LIFE*, three executions). Look at all the creative before
concluding. This is now rule 7 in AGENTS.md.

**Packaged it as `category-creative-scan`** (built with `skill-creator`). Nine coded
dimensions rather than the four in `competitor-comms-audit`: the four the user asked for —
`product_focus` (catches ecosystem convergence: EFU→Thrive/PRIMUS, Jubilee→Active,
IGI→Vitality), `audience_generation`, `content_type` (brand-building vs tactical vs PR),
and computed cadence — plus `who_is_in_frame` (the customer:executive ratio that diagnosed
State Life), `production_value`, `ecosystem_push`, `campaign_platform` and `language_lead`.

Five scripts, all verified end to end: `make_handles.py` (24-column workbook, 10 bound
vocabularies), `collect_ig.py`, `contact_sheets.py`, `analyze_scan.py`, `build_deck.py`.
Test cases in `evals/evals.json` — one of the three checks the skill *refuses* to answer a
share-of-voice question from a capped equal-N sample.

`sync-skills.sh` now also copies to `D:/Personal/Skills_aug2026/skills-main/skills/`.
AGENTS.md gains a Skills section and six evidence rules for competitive work.

Still open: market share (`MarketShare` empty), 25 `unclear` rows, no paid-channel data,
and `@pqftl_official` at 129 followers looks like a secondary account.

## 2026-08-26 — `brief-to-plan` skill; skills moved into the repo

Added **`brief-to-plan`** — parse a brief/RFP/tender/ToR into a statement of the ask and a
sequenced plan. It runs upstream of `competitor-comms-audit` and hands off to it.

- `scripts/extract_brief.py` — PDF (pdftotext preferred, pypdf fallback) and DOCX (body walk so
  tables come through). Prints chars/page and warns when a third of pages return footers only.
  Reports **cross-references with no matching heading** — sorted fewest-mentions-first, with the
  descriptive name the reference gives.
- `scripts/plan_from_register.py` — backward-pass CPM over `requirements.yaml`. Duration is
  `effort_days + lead_time_days`, separating our working time from time in someone else's hands.
  Emits start-today, already-late, critical path, gates, marks-per-day, and register.xlsx.
- `references/register-schema.md`, `references/extraction-gotchas.md`.

Two design decisions worth keeping: **sequence by latest possible start, not by marks** — a
2-mark bank certificate with a 4-day external wait outranks a 15-mark essay; and the model
fills `requirements.yaml` (judgment) while the script does the arithmetic (determinism).

Fixed a real modelling bug found on the SLIC data: items with a deadline of their own — a
clarification window, a site visit — were back-planning to the submission date and looking less
urgent than they are. Added a `due` override.

Verified on the actual tender: extractor pulled 44,339 chars across 37 pages and flagged
**Section VII → "Schedule of Requirements"** at the top of the unresolved list, which is the
genuine gap. Planner on `workspace/brief/requirements.yaml` (24 requirements) returns 8 days to
deadline, **1 overdue, 9 gates, critical-path slack −1 day** — the category audit needed to
start yesterday.

**Skills now live in the repo.** `.claude/skills/` is the source of truth (versioned, and it
triggers in-project); `.claude/skills/sync-skills.sh` copies both to `~/.claude/skills/` so they
also trigger from the UNICEF, PSO and future pitch folders. Edit in the project, then sync.

## 2026-08-26 — competitor audit scaffold + `competitor-comms-audit` skill

Built the competitive-analysis scaffold for the 15-mark Competitive Analysis criterion.

**Checked first:** Section VII "Schedule of Requirements" is referenced at ITA 6.8 but is *not
in the PDF*, and the PDS points to "section services and Lots" which lives on the EPADS portal.
So SLIC may have published a campaign brief we have not seen. Clarification window closes
**31 Aug** (ITA 7.1) — check EPADS Lots, and file a clarification asking which product line,
audience and objective the campaign submission should address, plus expected format/page limit.

**Packaged the method as a skill, not an agent.** Skills persist a repeatable method and load
into the current conversation; subagents are an execution strategy with their own context that
returns a report. A competitor comms audit recurs on every pitch, so it belongs in a skill —
with subagent fan-out written in as an optional collection step for >6 brands.

`~/.claude/skills/competitor-comms-audit/` (user-level, so it triggers from any pitch directory):

- `SKILL.md` — scoping (competitive set in four rings, not head-to-head), collection,
  coding, analysis, and the output structure. Central rule: if the analysis could be deleted
  without weakening the campaign concept, it has failed.
- `references/coding-schema.md` — per-ad schema. Claim territory (9 codes), emotional register
  (6), proof device, CTA channel, audience signal, objection handling. Includes inter-coder
  calibration and how to adapt the vocabulary per category.
- `references/sources.md` — per-channel collection guidance and failure modes. Notably: ad
  transparency libraries do **not** expose spend for commercial advertisers, so share of voice
  is counted as a presence proxy and never presented as spend.
- `scripts/make_capture_sheet.py` — generates a capture workbook with controlled vocabularies
  bound as dropdowns, so several coders work in parallel without drifting.
- `scripts/analyze_audit.py` — filled sheet → findings.md + CSVs: sample frame, claim×register
  territory map, per-brand claim profile, SOV-proxy vs share of market, proof/objection rates,
  and a coder-drift check. Rows without a `source_url` are dropped.

**SLIC instance** in `workspace/audit/`: `brands.csv` with 16 brands across four rings — State
Life (client); EFU, Jubilee, Adamjee, IGI, Askari, TPL (direct); five Takaful operators
including the separately-advertising window arms (adjacent); Postal Life, CDNS, Al Meezan, UBL
Funds (substitute). `capture.xlsx` generated and ready to fill.

Verified end to end on 81 synthetic rows: dropdowns bind, the no-source row is excluded, the
crosstab renders white space, and the drift check separates coders. Test artifacts deleted.

## 2026-08-26 — UNICEF pitch pack compared, SLIC form pre-filled

Added `data/UNICEF_PITCH/` — the UNICEF Pakistan ToR (Communication for Digital Fundraising,
1 Sep 2026 to 31 Dec 2027) and our submitted technical proposal v6 (Synite Digital (Pvt.) Ltd.,
10 Aug 2026). Compared both against the SLIC prequalification; findings in
`workspace/unicef_to_slic_reuse.md`.

**The material finding is an entity question.** SLIC ITA 22.1/23.3 count only the applicant's
own qualifications — parents, subsidiaries and affiliates are excluded — and the PDS caps JV
members at nil. The UNICEF proposal draws its scale from the Synergy Group ("four offices…
250+ professionals", "as part of the Synergy Group"), which was valid there and scores zero
here. Separately, nothing in the UNICEF pack evidences APNS, PBA, AAP or PID accreditation,
which SLIC treats as pass/fail plus 3 marks. The applicant entity must hold those in its own
name, which points away from Synite Digital (the digital practice, inc. 9 Feb 2018) and toward
whichever Group entity holds press/broadcast accreditation. Unresolved pending sight of the
certificates.

**Pre-filled `workspace/status.yaml`** from evidence in the pack, each note citing the proposal
section it came from; anything not evidenced there is left `missing`. Standing at 15.5/100 —
compliance 2/10, capability 13.5/20, campaign 0/70. Note this contradicts AGENTS.md rule 6 in
letter but not intent: the user supplied the source documents and asked for the form to be
filled from them. No answer was assumed — only transcribed with its source.

Reuse summary: the pack closes about half the paperwork (profile, five-year client list, staff
particulars and CVs, tax/registration annexures, Synergy Dentsu as the foreign affiliation) and
supplies a strategic spine for the 70-mark campaign — the trust/relevance/impact/ease/confidence
donor model maps onto life-insurance purchase barriers, and the State Bank of Pakistan and Polio
Eradication credentials fit a state-owned life insurer far better than the digital-fundraising
ROAS material does. Campaign concept and creative artworks (30 marks) transfer not at all.

## 2026-08-26 — initial build

**Scaffolded `slic-pq`** — a uv-managed Python package (`src/slic_pq`, Python 3.11+,
console script `slic-pq`) for the SLIC advertising-agency prequalification, EPADS ref
P78118.

Read the tender at `data/pq-bidding-document (2).pdf` (37 pages, issued 18 Aug 2026) and
encoded its requirements:

- `criteria.py` — 10 pass/fail eligibility documents (PQ p.21-22) and 16 scored technical
  criteria totalling 100 marks (PQ p.22-23), each carrying its source page. Criteria are
  grouped into three buckets: corporate compliance (10), agency capability & scale (20),
  campaign submission (70).
- `extract.py` — PDF text extraction and section slicing into `invitation`,
  `instructions`, `pds`, `eligibility`, `evaluation`, `annexures`, `forms`, plus a
  `key_dates()` reader for the PDS dates.
- `readiness.py` — `workspace/status.yaml` self-assessment, validated and scored against
  the 50-mark pass bar; `ready`/`partial`/`missing`/`na` map to 1.0/0.5/0.0/0.0.
- `reports.py` — deliverables: submission checklist as Markdown and Excel (summary,
  eligibility, technical scoring, gaps sheets) and the Annexure I/II/III forms as Word.
- `cli.py` — commands `info`, `criteria`, `parse`, `init`, `score`, `gaps`, `build`.

**Two extraction findings worth keeping:**

1. pypdf recovers almost nothing from this file — the body text sits in XForm objects it
   cannot decode, so pages 3-20 return only their footer (~86 characters each), 14.9k
   characters total against ~44k of real text. Switched extraction to prefer poppler's
   `pdftotext -layout`, with pypdf retained as a labelled fallback.
2. Section headings recur in the table of contents and inside ITA cross-references.
   Section slicing now starts after `INVITATION FOR PRE-QUALIFICATION` to skip the TOC,
   and the PDS anchors on its own opening sentence rather than the heading, which ITA 6.3
   also uses. Before the fix the `instructions` section came out at 6,271 characters;
   after, 24,176.

**Verified end to end** on a sample assessment: `init` → edit → `score` (79/100, above the
50 pass mark) → `gaps` → `build` produced all three deliverables. Sample outputs and the
sample status file were then cleared; `workspace/status.yaml` ships as a blank template.

Deliverables are date-stamped and never overwritten — same-day rebuilds get `-v2`, `-v3`.
