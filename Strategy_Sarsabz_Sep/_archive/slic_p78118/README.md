# slic-pq

Bid-readiness toolkit for the **State Life Insurance Corporation of Pakistan (SLIC)
prequalification of advertising agencies**, EPADS reference **P78118**, engagement
term 2026-28.

It reads the prequalification PDF in `data/`, holds the eligibility and evaluation
criteria as traceable data, scores your agency's self-assessment against the 100
technical marks, and builds the submission checklist and Annexure forms.

**Submission deadline: Thursday 3 September 2026, 11:00 AM, through EPADS v2.0 only.**
Manual submission is not entertained.

## Install

```bash
uv sync
```

Text extraction prefers poppler's `pdftotext` (already on PATH in this environment).
The EPADS PDF stores its body text in XForm objects that pypdf cannot decode, so the
pypdf fallback returns page footers only -- keep `pdftotext` available.

## Use

```bash
uv run slic-pq info                      # tender at a glance: dates, method, marks
uv run slic-pq criteria                  # eligibility documents + scored criteria
uv run slic-pq criteria --json           # same, as JSON
uv run slic-pq parse                     # slice the PDF into named sections
uv run slic-pq parse --section evaluation
uv run slic-pq parse --out workspace/pq.txt

uv run slic-pq init --agency "Your Agency (Pvt) Ltd"   # create workspace/status.yaml
# ... edit workspace/status.yaml: ready | partial | missing | na ...
uv run slic-pq score                     # marks earned vs the 50-mark pass bar
uv run slic-pq gaps                      # what is costing marks, worst first
uv run slic-pq build                     # checklist (md + xlsx) + Annexures (docx)

uv run slic-pq doc workspace/concept/the_concept.md     --stem SLIC_P78118_campaign_concept   # render any written section as a .docx
```

`ready` scores full marks, `partial` scores half, `missing` and `na` score nothing.
The half-credit rule is our own planning heuristic, not SLIC's -- the PQ document does
not publish a sub-scoring rubric.

## Competitive analysis: the category scan

The 70-mark campaign submission opens with Competitive Analysis (15 marks). That work is
a repeatable method packaged as skills, not one-off scripts — see **Skills** below.

Everything lives in `workspace/audit/`. Run in this order:

```bash
S=~/.claude/skills/category-creative-scan/scripts

# 1. resolve Instagram handles by SEARCH (never guess them), then collect.
#    --depth 24 is the logged-out ceiling; --enrich opens each post for its
#    true date and full caption, which the profile grid does not carry.
uv run python $S/collect_ig.py workspace/audit/handles.csv     --out workspace/audit/deep --depth 24 --enrich

# 2. contact sheets, then LOOK at them - coding is from the artwork, not the caption
uv run --with pillow python $S/contact_sheets.py workspace/audit/deep

# 3. carry provenance into the capture sheet, ready to code
uv run --with openpyxl python $S/seed_capture.py workspace/audit/deep

# 4. apply codes from a CSV (validated against the vocabulary before anything is written)
uv run --with openpyxl python $S/apply_codes.py workspace/audit/codes_deep.csv     --workspace workspace/audit/deep

# 5. analysis -> scan_findings.md + seven CSVs
uv run --with openpyxl python $S/analyze_scan.py workspace/audit/deep/capture.xlsx     --out workspace/audit/deep

# 6. choose which posts appear on each slide, by a stated rule
uv run --with openpyxl python $S/select_for_slides.py workspace/audit/deep --per-brand 12
uv run --with pillow python $S/contact_sheets.py workspace/audit/deep     --selection workspace/audit/deep/slide_selection.csv

# 7. the two deliverables
uv run --with python-pptx --with pyyaml --with openpyxl --with pillow     python $S/build_deck.py workspace/audit/deep     --title "SLIC P78118 category audit" --out workspace/out
uv run --with python-docx --with openpyxl     python workspace/audit/build_diagnostic_doc.py
```

The project-specific coding rules live in `workspace/audit/build_codes_deep.py` and
`build_claims_deep.py` — they match on caption text rather than transcribed post ids,
because shortcodes contain I/l/1 lookalikes and hand-copying them fails silently.

**Two deliverables, deliberately.** The deck carries the argument; the Word document
carries the evidence. SLIC scores every criterion `(Qualitative)(Doc Required)`, so each
scored line needs an uploaded document and a slide pack alone is a thin answer.

## The campaign submission

70 of the 100 marks are the speculative campaign, scored in five parts. Each is
`(Qualitative)(Doc Required)` — a written section **and** an uploaded document.

| Criterion | Marks | Source | Status |
|---|---:|---|---|
| Campaign Concept | 15 | `workspace/concept/the_concept.md` | drafted |
| Competitive Analysis | 15 | `workspace/audit/the_finding.md` | drafted |
| Communication Strategy | 15 | `workspace/strategy/the_strategy.md` | drafted |
| Creative Artworks | 15 | — | not started |
| Proposed Media Mix | 10 | — | not started |

Each written section is Markdown first, rendered to Word by the same command, so the
working file and the submitted file cannot drift:

```bash
uv run slic-pq doc workspace/concept/the_concept.md  --stem SLIC_P78118_campaign_concept
uv run slic-pq doc workspace/strategy/the_strategy.md --stem SLIC_P78118_communication_strategy
```

The Competitive Analysis document is built differently — it composes its tables live from
the coded dataset rather than from prose, so the figures cannot drift from the sheet:

```bash
uv run --with python-docx --with openpyxl python workspace/audit/build_diagnostic_doc.py
```

**Two rendering rules** for anything written for `slic-pq doc`: list items and blockquotes
may wrap across lines (continuations are joined), but **nested emphasis is not supported** —
write `**BOLD** — *italic*`, never `**BOLD — *italic***`. Both rules exist because both were
shipped as bugs first.

## Skills

`.claude/skills/` is the source of truth, versioned in git. `sync-skills.sh` copies each
skill to `~/.claude/skills/` (so it triggers from any pitch folder) and to the shared
library. Edit here, then sync — never edit the copies.

```bash
bash .claude/skills/sync-skills.sh
```

| Skill | What it does |
|---|---|
| `brief-to-plan` | Brief / RFP / tender → statement of the ask + backward-pass CPM plan |
| `category-creative-scan` | Collect a category's live social creative, code it on nine dimensions, brand-wise deck |
| `competitor-comms-audit` | The strategic method: rings, message territory, how to write the finding |
| `campaign-concept` | Audit finding → core idea, proposition, line, executions, rationale |
| `communication-strategy` | Objective, audience, approach, message architecture, journey |

They chain: **brief-to-plan → category-creative-scan → competitor-comms-audit →
campaign-concept → communication-strategy**.

## Layout

```
data/                  the prequalification PDF as issued
src/slic_pq/
  criteria.py          eligibility + evaluation criteria, each with its PQ page
  extract.py           PDF text extraction and section slicing
  readiness.py         status.yaml load/validate and scoring
  reports.py           Markdown, Excel and Word deliverables
  submission_doc.py    Markdown -> Word renderer for written submission sections
  cli.py               typer command line
workspace/status.yaml  your self-assessment (edit this)
workspace/brief/       the ask, the requirement register, the dated plan
workspace/audit/       the category scan: handles, capture.xlsx, media/, contact/
  deep/                the 288-post read - capture sheet, findings, contact sheets
  the_finding.md       Competitive Analysis (15 marks)
workspace/concept/     the_concept.md - Campaign Concept (15 marks)
workspace/strategy/    the_strategy.md - Communication Strategy (15 marks)
workspace/out/         generated deliverables, date-stamped, never overwritten
```

Deliverables are versioned: a rebuild on the same day writes `-v2`, `-v3` rather than
overwriting the previous file.

See `AGENTS.md` for working conventions and `HISTORY.md` for the change log.
