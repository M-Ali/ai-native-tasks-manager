# AGENTS.md — working conventions for this repo

## What this project is

`slic-pq` supports one live bid: the **SLIC prequalification of advertising agencies**,
EPADS reference **P78118**, term 2026-28. The source of truth is the tender document at
`data/pq-bidding-document (2).pdf`, issued 18 August 2026.

Hard dates (from the Prequalification Data Sheet):

| Event | When |
|---|---|
| Clarification / pre-application meeting | Monday 31 August 2026 |
| Application submission deadline | **Thursday 3 September 2026, 11:00 AM** |
| Opening of applications | Thursday 3 September 2026, 11:30 AM |

Everything moves through EPADS v2.0 — clarifications, addenda, submission, grievances.
Manual submission is rejected outright.

## Ground rules

1. **The PDF is the authority.** Every criterion in `criteria.py` carries the PQ page it
   was read from. If you change a criterion, re-read that page first and update the page
   number with it. Do not add a requirement the document does not state.
2. **Extraction uses `pdftotext`.** The EPADS PDF hides its body text in XForm objects;
   pypdf returns footers only (~86 chars/page). `extract.load()` prefers poppler and
   falls back to pypdf with the engine recorded on the `Document`. Do not switch the
   default back to pypdf.
3. **Never overwrite a deliverable.** `reports._versioned()` date-stamps output and adds
   `-v2`, `-v3` on same-day rebuilds. Keep it that way — prior versions of a bid document
   are evidence of what was reviewed and when.
4. **Half credit is ours, not SLIC's.** The `partial = 0.5` rule in `readiness.CREDIT` is
   a planning heuristic. SLIC publishes marks per criterion and a 50-mark pass bar, no
   sub-rubric. Say so wherever a score is presented.
5. **Only the applicant's own qualifications count.** ITA 22.1 and 23.3 exclude
   subcontractors, subsidiaries, parents and affiliates from evaluation. Do not model
   borrowed capability into the score.
6. **`workspace/status.yaml` is user data.** Regenerate it only on explicit `--force`.
   Do not fill it with assumed answers about the agency.

## Layout

```
data/                  tender documents as issued (read-only)
src/slic_pq/
  criteria.py          eligibility + scored criteria, page-traced
  extract.py           pdftotext/pypdf extraction, section slicing
  readiness.py         status.yaml load, validation, scoring
  reports.py           Markdown / Excel / Word deliverables
  cli.py               typer entry point (`slic-pq`)
workspace/status.yaml  the agency self-assessment
workspace/out/         generated, date-stamped deliverables
```

## Conventions

- Python 3.11+, managed with `uv`. Add dependencies with `uv add`, never by hand-editing
  `pyproject.toml` without a matching `uv sync`.
- Run everything through `uv run slic-pq ...`.
- Dataclasses over dicts for anything derived from the tender; frozen where it is
  tender-stated fact.
- Keep the CLI thin: commands assemble and print, logic lives in the modules.
- No network calls. This is an offline bid-prep tool.

## Where the marks are

| Bucket | Marks | Character |
|---|---:|---|
| Corporate compliance (SECP, FBR, APNS/PBA/PID, clean-record declaration) | 10 | Documentary, binary |
| Agency capability & scale (experience, clients, revenue, offices, headcount, staff) | 20 | Documentary, thresholded |
| Campaign submission (concept, competitive analysis, strategy, artworks, media mix) | 70 | Subjective, must be created |

70 of 100 marks are the speculative campaign. Compliance paperwork is necessary but
cannot win this — plan effort accordingly.

## Skills

`.claude/skills/` is the source of truth, versioned in git. `sync-skills.sh` copies each
skill to `~/.claude/skills/` (so it triggers from other pitch folders) and to the shared
library at `D:/Personal/Skills_aug2026/skills-main/skills/`. Edit here, then sync — never
edit the copies.

| Skill | Does |
|---|---|
| `brief-to-plan` | Brief/RFP/tender -> statement of the ask + backward-pass CPM plan |
| `category-creative-scan` | Collect a category's live social creative, code it, brand-wise deck |
| `competitor-comms-audit` | The strategic method: rings, territory map, the finding |
| `campaign-concept` | Audit finding -> core idea, proposition, line, executions, rationale |
| `communication-strategy` | Objective, audience, approach, message architecture, journey |

They chain: **brief-to-plan -> category-creative-scan -> competitor-comms-audit ->
campaign-concept -> communication-strategy**.

## Evidence rules for competitive work

These were learned the expensive way on this bid. They apply to any category scan.

7. **Code from the artwork, not the caption.** A caption states intent; the image shows
   execution, casting, budget and production value. Reading four of twelve images on this
   scan produced a confidently wrong conclusion about the client's own advertising that
   the full grid overturned. Look at everything before concluding anything.
8. **`unclear` is a valid code; a guessed code is not.** 25 of 144 rows in the SLIC scan
   are `unclear`. That is honest data. Inferring a code from a brand's reputation is not.
9. **Never present share of voice from a capped sample.** Logged-out collection caps at
   ~12 posts per brand, which makes every brand an equal share by construction. The number
   is an artefact. State the cap in section 1 rather than burying it.
10. **Organic is not paid.** Owned-channel posts say nothing about spend, reach or media
    weight. Only the Meta Ad Library shows paid creative, and it shows no spend for
    commercial advertisers. Never present an estimated spend figure as measured.
11. **Search for handles; never guess them.** A guessed handle on this scan landed on a
    personal account with 30 followers. Verify against follower count and bio.
12. **Absence is a finding.** Two substitutes had no Instagram at all. Record it; it says
    where the category actually competes.

## Writing the scored sections

Each of the five campaign criteria is `(Qualitative)(Doc Required)`: a written section and
an uploaded document. Conventions that keep them defensible:

13. **Markdown is the source; Word is a render.** Write the section as `.md`, render with
    `slic-pq doc`. Never edit the `.docx` — the next render silently discards it, and the
    working file and the submitted file drift apart.
14. **Every figure re-verified against the sheet before it is written**, not after. Six
    figures were checked before the concept was drafted; one error (`fear` register) was
    caught in the finding this way.
15. **State the assumption on the page.** Section VII was never issued and the clarification
    window closed, so product line, audience and objective are ours. A labelled assumption
    reads as judgement; an unstated one reads as an oversight.
16. **Say what the evidence cannot show.** The audit is a supply-side read of what the
    category says, not consumer research. The strategy's barrier diagnosis is an inference
    and says so. A section that overstates its evidence is easy to dismantle in the room.
17. **Distinguish a data request from an assumption.** Pillar 1 rests on complaints
    resolved, which State Life publishes. A claims-paid figure would be stronger — that is
    a question for the client, not a number to invent.
18. **Check dates and heritage claims before they reach artwork.** State Life is 54 years
    old (1972), not 60. A checkable number that is wrong costs more than a smaller one that
    is right.

## Change log

Record every material change in `HISTORY.md`, newest first, dated.
