# sarsabz-pitch

Working toolkit for the **Fatima Fertilizer creative-agency pitch 2026** (Sarsabz, Bubber Sher, Pakarab CAN).
Brief: `data/Creative Pitch Deck (1).pdf` (32 pages, 13 of them image-only). Farmer comments supplied by the client team:
`data/COMMENTS.xlsx`.

---

## Start here (state at end of 15 September 2026)

**Case 2 is built.** Open **`workspace/out/Sarsabz_case2_deck-v4.pptx`** (36 slides, internal working draft):

| Slides | Section | Source of truth |
|---|---|---|
| 1 | Cover | `comments/build_consumer_section.py` |
| 2-6 | What farmers say (1,433 comments) | `workspace/case2_audit/comments/analysis.md` |
| 7-16 | How the category talks (Sarsabz vs Engro, 200 coded uploads) | `scan/scan_findings.md`, `scan/deck/reads.yaml` |
| 17-21 | Brand Meaning Ladder | `scan/ladder_spec.json`, `scan/ladder_levels.csv` |
| 22-28 | The Big Idea: *Sarsabz proves itself on a field like yours* | `big_idea.md`, `big_idea_spec.json` |
| 29-36 | The campaign: **Dus Feesad Aur. Khet Gawah Hai.** | `concept/concept.yaml`, `concept/the_concept.md` |

Paths above are under `workspace/case2_audit/`. The full argument in prose is `workspace/case2_audit/the_audit.md`.
**Case 1 (Salam Kissan 2026) is not started.**

### Next, in priority order

1. **Ask the client** (none of these are in the brief):
   - Pitch date and budget.
   - The demo-plot dataset behind the 10% claim.
   - Ki Jeet winners' input records.
   - Which YouTube channels `COMMENTS.xlsx` came from, and whether any were Fatima-sponsored.
   - Competitor names and method for the brand-health study (p29).
   - Full list: `uv run sarsabz-pitch info`.
2. **Case 1, Salam Kissan 2026.** A 360° campaign reaching rural, urban and young audiences, beyond one day (brief p12). The
   `big-idea` method needs its own consumer evidence here: the farmer comments don't cover urban or young audiences.
3. **Bubber Sher.** The brief expects "strategic vision for Sarsabz and Baber Sher" (p14), and neither case covers it.
4. **Client-facing Case 2 deck.** A lower-volume version: no file names or source notes, hostile verbatims made neutral. Both
   `brand-laddering` and `big-idea-slides` describe how.
5. **Evidence gaps** (audit §6):
   - FFC and Engro paid ads, from the Meta Ad Library
   - Sarsabz and Engro Facebook and TikTok content
   - Comments on Sarsabz's own 10% and Ki Jeet films
   - Fatima annual report 2025
   - NFDC Fertilizer Review 2024-25
6. **Agency sections of the 90-minute pitch** (intro, team, awards, case studies; brief p11-12). Agency facts, not ours to invent.
7. **Housekeeping decisions for the user:**
   - Whether to sync the patched `build_deck.py` to `~/.claude/skills` (see AGENTS.md).
   - Whether to turn the consumer and concept builders into skills.

### Other deliverables in `workspace/out/` (latest versions)

| File | What |
|---|---|
| `Sarsabz_case2_brand_audit_draft_2026-09-15-v5.docx` | The audit, rendered from `the_audit.md` |
| `Sarsabz_farmer_conversation_analysis_2026-09-15.docx` | The comment analysis |
| `Sarsabz_big_idea_sheet_2026-09-15-v2.docx` | One-page Big Idea |
| `Sarsabz_campaign_concept_2026-09-15.docx` | Campaign rationale |
| `Sarsabz_brand_ladder_section-v2.pptx`, `Sarsabz_big_idea_section-v2.pptx` | Standalone sections (superseded by the combined deck) |

Earlier `-vN` files are kept deliberately. Never delete or overwrite a deliverable.

---

## Install and use

```bash
uv sync
uv run sarsabz-pitch info        # brief at a glance + open questions for the client
uv run sarsabz-pitch case2       # every Case 2 ask, with the brief page it came from
uv run sarsabz-pitch facts       # brand facts and brand-health numbers
uv run sarsabz-pitch parse       # full.txt + PNGs of the image-only pages -> workspace/brief/
uv run sarsabz-pitch doc workspace/case2_audit/the_audit.md --stem Sarsabz_case2_brand_audit_draft
```

## Rebuild the deck

The deck is always built, never hand-edited. Put intermediates in a scratch folder, never in `workspace/out/`, because
`build_deck.py` overwrites an explicit `.pptx` path. The last step writes the next free `workspace/out/Sarsabz_case2_deck-vN.pptx`.

```bash
S=<scratch folder>
# spec copies whose "out" points at $S (originals stay untouched)
python -c "import json,sys; S=sys.argv[1]
for src,dst,out in [('workspace/case2_audit/scan/ladder_spec.json','ladder.json','03_with_ladder.pptx'),
                    ('workspace/case2_audit/big_idea_spec.json','big_idea.json','04_with_big_idea.pptx')]:
    d=json.load(open(src,encoding='utf-8')); d['out']=f'{S}/{out}'; json.dump(d,open(f'{S}/{dst}','w',encoding='utf-8'),ensure_ascii=False)" "$S"

uv run --no-project --with python-pptx --with pyyaml --with openpyxl python .claude/skills/category-creative-scan/scripts/build_deck.py \
    workspace/case2_audit/scan/deck --title "How the category talks" --client "Sarsabz - Case 2" --out "$S/01_scan.pptx"
uv run --no-project --with python-pptx python workspace/case2_audit/comments/build_consumer_section.py \
    --base "$S/01_scan.pptx" --cover --front --out "$S/02_consumer_scan.pptx"
uv run --no-project --with python-pptx --with pandas python ~/.claude/skills/brand-laddering/scripts/build_ladder_section.py \
    "$S/ladder.json" --levels workspace/case2_audit/scan/ladder_levels.csv --base "$S/02_consumer_scan.pptx"
uv run --no-project --with python-pptx python ~/.claude/skills/big-idea-slides/scripts/build_big_idea_section.py \
    "$S/big_idea.json" --base "$S/03_with_ladder.pptx"
uv run --no-project --with python-pptx --with pyyaml python workspace/case2_audit/concept/build_concept_section.py \
    --base "$S/04_with_big_idea.pptx" --out workspace/out/Sarsabz_case2_deck.pptx
```

Then render and look at every changed slide. PowerPoint COM in PowerShell, late-bound: `SaveAs(<dir>, 18)` writes one PNG per slide.
The recipe is in AGENTS.md.

## Regenerate the evidence (only if the data changes)

Run from `workspace/case2_audit/`. YouTube serves a bot check after a few hundred requests, so space out runs.

| Evidence | Commands |
|---|---|
| Sarsabz owned YouTube | `owned/`: yt-dlp flat listings → `build_inventory.py` → `enrich_dates.py` → `code_window.py` → `analyze_window.py` |
| Engro owned YouTube | `scan/engro/`: same scripts, its own `code_window.py` |
| Comparison | `scan/compare.py` → `scan_findings.md` |
| Ladder levels | `scan/to_levels.py`, then `~/.claude/skills/brand-laddering/scripts/compute_levels.py levels.csv --out ladder_levels.csv` |
| Scan deck inputs | `scan/deck_prep.py` (capture.xlsx, handles, profiles, contact sheets); `scan/deck/reads.yaml` is hand-authored |
| Farmer comments | `comment-analysis` skill: `load_comments.py data/COMMENTS.xlsx` → `code_comments.py --lexicon comments/lexicon.yaml` |

Presence sweep: `category-creative-scan/scripts/sweep_presence.py scan/candidates.csv`. Primary documents: `sources/`.

## Layout

```
data/                           the brief, COMMENTS.xlsx, reference/ (Levels of Branding)
src/sarsabz_pitch/              brief.py (page-traced asks), extract.py, submission_doc.py, cli.py
workspace/brief/                full.txt + page images
workspace/case2_audit/
  the_audit.md                  the argument
  big_idea.md, big_idea_spec.json
  concept/                      concept.yaml, the_concept.md, build_concept_section.py
  comments/                     analysis.md, lexicon.yaml, coded CSVs, build_consumer_section.py
  owned/                        Sarsabz channel inventory, contact sheets, codes, findings
  scan/                         presence, engro/, compare.py, levels, ladder spec, deck/ (reads.yaml + inputs)
  sources/                      CCP, VIS, PACRA, BMA, Economic Survey, Dawn - downloaded + extracted text
workspace/out/                  every rendered deliverable, versioned
_archive/                       SLIC and PTCL work this repo was copied from; skill drafts; misfiled copies. Read-only
.claude/skills/                 project skill copies (see AGENTS.md for which live where)
```

See `AGENTS.md` for conventions and hard-won rules, and `HISTORY.md` for what happened when.
