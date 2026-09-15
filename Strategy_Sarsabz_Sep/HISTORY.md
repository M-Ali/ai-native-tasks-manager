# HISTORY

Newest first. The history of the SLIC and PTCL work this repo was copied from is in `_archive/slic_p78118/HISTORY.md`.

> **Resuming?** Read `README.md` → "Start here". At the end of 15 Sep 2026, Case 2 is complete as
> `workspace/out/Sarsabz_case2_deck-v4.pptx` (36 slides), Case 1 is not started, and the pitch date and budget are still unknown.

## 2026-09-15 (after midnight) — campaign concept added to the deck

**Concept** via the `campaign-concept` skill (`workspace/case2_audit/concept/`):
- `concept.yaml` holds the chain; `the_concept.md` holds the rationale in the skill's six sections, rendered to `Sarsabz_campaign_concept_2026-09-15.docx`.
- **Line: "Dus Feesad Aur. Khet Gawah Hai."** (ten percent more, the field is the witness). It revives an owned, dormant line found in
  the channel archive ("Dus Feesad Aur!" wheat 2018, 370K views; rice TVC 2021, 1.8M) and adds a proof sign-off.
- **Mechanic:** Do Khet (paired plots), Kaanta Din (independent harvest weigh-in), Nuskha (published recipe).
- **Lines considered:** "Pehle Dekho, Phir Daalo"; reviving "Khaad Muft He Samjho" (walks into Engro's cost fight); keeping the brief's
  "10% se bhi ziyada" wording.
- **Asks of Fatima:** publish every plot result including under 10%, clear the 10% claim, agronomy signs the Nuskha, move the
  sticker code from the prize draw to verification, appoint an independent adjudicator.

**Deck:** `workspace/out/Sarsabz_case2_deck-v3.pptx`, 36 slides.
- The concept section (8 slides) was appended by `concept/build_concept_section.py`, which refuses to build unless every archive
  title and brief phrase it cites is found and a month-twelve entry exists.
- v2 was the append; v3 adds "5 The campaign" to the cover agenda. The consumer builder's cover text is updated to match.
- v1 (28 slides) is kept.

**No concept slide skill exists** in any skill folder. The builder is project-level for now.

**v4 is the current deck** (`Sarsabz_case2_deck-v4.pptx`, 36 slides), rebuilt end to end from source rather than hand-patched. The
v3 render check found the line slide's 30 pt headline wrapping onto the translation, and over-tall calendar boxes. Both were fixed in
`build_concept_section.py`, and the cover's five-section agenda now comes from the consumer builder. v1-v3 are kept.

## 2026-09-15 (late night) — comms scan and consumer slides; combined Case 2 deck

**Combined deck:** `workspace/out/Sarsabz_case2_deck.pptx`, 28 slides. Order: cover → What farmers say (consumer, 5) → How the
category talks (scan, 10) → Brand Meaning Ladder (5) → The Big Idea (7).

- **Comms scan section:** the user's `category-creative-scan/scripts/build_deck.py`. The coded YouTube read was shaped into its
  inputs by `scan/deck_prep.py` (Ads sheet, handles, profiles, a 25-most-viewed contact sheet per brand) with `scan/deck/reads.yaml`.
- **Consumer section:** no slide skill exists, so `comments/build_consumer_section.py` was written in the house style. Numbers are read
  live from the coded CSVs, and it refuses to build if any quote isn't verbatim in `comments.csv` (translations flagged).
- **Merge:** the consumer builder inserts the cover and section in front of the scan deck (`--base --front`); the ladder and Big Idea
  builders append with `--base`. Intermediates were kept in the scratchpad; the combined spec copies leave the original specs untouched.

**Two hardcoded category leftovers fixed in the project copy of `build_deck.py`** (original backed up to
`_archive/draft_skills/build_deck.py.before-2026-09-15`; not yet synced to `~/.claude/skills`):
1. The fourth headline stat was insurance-specific ("ads cast for over-60s… a category selling retirement and legacy"). It's now
   the objection rate, which applies to any category.
2. The "What we looked at" slide had an Instagram method hardcoded ("most recent posts… coded from the image… organic, no spend"),
   which misstated a YouTube read of mostly bought views. It's now overridable per scan through a `_method` key in `reads.yaml`,
   with the old text as the default. Caught on the render check.

## 2026-09-15 (night) — the Big Idea

**Skills.** The user asked for a new "The Big Idea" skill, but checking first showed two already existed. `big-idea-slides` was already
installed; `big-idea` (the method) was only in `D:\Personal\Strategy_PTCL_Sep\.claude\skills\` and was copied, never overwritten,
to `~/.claude/skills/` and the shared library (checksums match). No new skill was written.

**Sarsabz Big Idea** via `big-idea` (method) and `big-idea-slides` (builder, unchanged):
- Sheet: `workspace/case2_audit/big_idea.md`, rendered to `Sarsabz_big_idea_sheet_2026-09-15.docx`.
- Deck: `workspace/out/Sarsabz_big_idea_section-v2.pptx` (7 slides; v1 kept).
- **Idea:** *Sarsabz proves itself on a field like yours.* Platform: THE PROOF, NOT THE PROMISE. Alternatives: the farmer's adviser
  (Engro's territory), saluting the farmer (Sarsabz's current level 5, kept for Salam Kissan).
- Every quote in the spec was checked by script against `brief/full.txt`, `comments.csv` and the YouTube descriptions before building.
  One non-verbatim brief quote was corrected ("TOM awareness").
- Visual check of v1 found two claims my own data contradicted: "the only NP and CAN" (Engro has 6% of NP) and "the only differentiated
  products" (Engro's zinc-coated urea). Both corrected in v2 and in the sheet.

## 2026-09-15 (evening) — competitor scan; brand ladder deck

**Presence sweep** (`workspace/case2_audit/scan/presence.csv`), with handles found by search, not guessed. Sarsabz TikTok 1.2M, YouTube 396K,
Facebook 274K; Engro 157K / 81K / 152K; FFC 71K Facebook with no official YouTube; Bubber Sher and Agritech with no official channel. One
guessed-looking hit ("Bubber Sher" Facebook) turned out to be a rock band and was excluded.

**Engro coded** (`scan/engro/`): 134 uploads inventoried, the 110 inside the 24-month window coded on the same schema after reading every
thumbnail. Two claim values were added (`cost_saving`, `authenticity`), both checked against Sarsabz's full channel first: Sarsabz
used both in 2017-22, then dropped them. `scan/compare.py` → `scan_findings.md`. Headline: Engro answers a farmer objection in 58% of
uploads, Sarsabz in 6%.

**Brand ladder deck** built with the user's existing `brand-laddering` skill (its `compute_levels.py` and `build_ladder_section.py`,
unchanged): `workspace/out/Sarsabz_brand_ladder_section.pptx` (v1) and `-v2` (wording fixes). The builder's consistency check refused
the first spec: brand labels were typed, the target row was missing and the output path was wrong. All fixed in the spec, not the skill.
Slides rendered through PowerPoint (late-bound COM; the early-bound type library is broken on this machine) and checked visually: no overflow.

**Skill note:** a `brand-laddering` skill already existed in `~/.claude/skills/` and the shared library when a new one was requested. The
duplicate draft was moved to `_archive/draft_skills/`, and `brand-laddering` was removed from `sync-skills.sh`, which would have `rm -rf`'d
the user's version.

## 2026-09-15 (later) — farmer conversation analysis; a misfiled input corrected

**Misfile, corrected.** During the repurposing, `data/COMMENTS.xlsx` was moved to `_archive/ptcl_telco/data/` on the
assumption that it was the old PTCL comment file of the same name. It wasn't: it was the Sarsabz farmer comments the user
had just placed (modified 16:45 that day). The user re-placed the file; the two copies were verified identical cell by cell
(1,531 rows), and the misfiled copy moved to `_archive/misfiled_2026-09-15/`. Rule added to AGENTS.md: open a file before
moving it; a matching filename is not a matching file.

**Farmer conversation analysis** (`workspace/case2_audit/comments/`, rendered `Sarsabz_farmer_conversation_analysis_2026-09-15.docx`)
via the `comment-analysis` skill: 1,433 comments, fertilizer lexicon written after reading the full file, 54.1% coverage.
One false positive found and removed (English "can" matching CAN: 254 → 193). Every quoted verbatim checked against the corpus.
Headline: farmers ask how to combine NP and CAN with what they already use (220 mixing comments), not whether they work;
Sarsabz is named in 15 comments against CAN 193 and NP 147; Engro Zabardast is the only brand named routinely.

**Audit updated to v3 render:** new §3.7 (demand side), sharpened finding, Ki Jeet caveat on progressive-farmer yields,
answers table and client questions revised.

## 2026-09-15 — repurposed for Fatima Fertilizer; Case 2 audit started

**Repo repurposed.** The folder was a copy of the SLIC P78118 bid toolkit (`slic-pq`) and also carried PTCL
fixed-broadband work. Nothing was deleted:

- `workspace/*` (SLIC audit, brief, concept, strategy) → `_archive/slic_p78118/workspace/`
- `src/slic_pq/`, old README/AGENTS/HISTORY/pyproject/uv.lock → `_archive/slic_p78118/`
- `workspace/scan_telco`, `workspace/comments_test`, `data/COMMENTS.xlsx`, broadband deck → `_archive/ptcl_telco/`
- Chery SUV analysis → `_archive/other_clients/`; `Levels of Branding.docx` → `data/reference/`

**New package `sarsabz-pitch`**: `brief.py` (page-traced asks, brand facts, awareness data, open questions),
`extract.py` (pypdf + PyMuPDF rendering of the 13 image-only pages; poppler no longer needed),
`submission_doc.py` (Markdown→Word renderer carried over, Fatima indigo / Sarsabz green), and CLI commands
`info`, `case2`, `facts`, `parse`, `doc`.

**Skills checked for hardcoding.** The scripts carry no client logic. The only category-specific parts are the
*default* claim vocabularies (financial services, overridable with `--claims`) and PTCL/telco examples in lesson
notes. Left as-is; a fertilizer vocabulary is passed per project instead.

**Case 2 audit, draft v0** (`workspace/case2_audit/the_audit.md`, rendered `Sarsabz_case2_brand_audit_draft_2026-09-15-v2.docx`):

- Category structure from primary sources: CCP draft study Table 11 (checked against the page image because Fatima's
  CAN cell is blank), VIS, PACRA, BMA, Economic Survey.
- Owned YouTube: 401 uploads inventoried with true publish dates; the 90 from the last 24 months were coded on the
  scan schema after reading every thumbnail. A `sponsorship` content type was added and documented.
- Sarsabz Ki Jeet checked against independent reporting: usable as proof, but it overclaims in two ways
  ("proved 10%", and "all" winners).

**Not done:** competitor creative scan, Facebook/TikTok presence, farmer comments (YouTube bot-walled after ~420
requests), Fatima annual report, NFDC 2024-25 review.
