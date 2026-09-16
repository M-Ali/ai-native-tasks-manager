# sarsabz-pitch

Working toolkit for the **Fatima Fertilizer creative-agency pitch 2026** (Sarsabz, Bubber Sher, Pakarab CAN).
Brief: `data/Creative Pitch Deck (1).pdf` (32 pages, 13 of them image-only). Farmer comments supplied by the client team:
`data/COMMENTS.xlsx`.

---

## Start here (state at end of 16 September 2026)

**Both cases are built as internal working drafts.** Everything is committed on branch `sarsabz-pitch` (last commit `6b6eeeb`).
**Nothing is pushed:** both remotes are public, and the branch holds the client brief and downloaded comments.

### Case 1: Salam Kissan 2026 → `workspace/out/Sarsabz_case1_deck-v3.pptx` (19 slides)

Built under the user's rule **"no made up data, no assumptions"**.

| Slides | Section | Source of truth (under `workspace/case1_salam_kissan/`) |
|---|---|---|
| 1-2 | Cover; what the brief asks (p12) | `src/sarsabz_pitch/brief.py` (`uv run sarsabz-pitch case1`) |
| 3 | Seven years on YouTube (76 uploads coded) | `archive/archive_findings.md`, `archive/archive_codes.csv` |
| 4-5 | What commenters say (1,572 comments, all read) | `comments/comment_findings.md`, `comments/analyse.py` → `comment_codes.csv` |
| 6 | Who else marks the day (YouTube scan) | `scan/scan_findings.md`, `scan/others.csv` |
| 7-13 | The Big Idea: *the farmer answers Pakistan's salute in his own voice, through his own family* | `big_idea.md`, `deck/big_idea_spec.json` |
| 14-19 | The campaign: **"Salam Kissan. Salam Pakistan."** (proposed), mechanic **Kissan ka Jawab** | `the_concept.md` |

- **Summary of all evidence:** `workspace/case1_salam_kissan/evidence.md`.
- **Rebuild:** `uv run --no-project --with python-pptx python workspace/case1_salam_kissan/deck/build_case1_deck.py`. It writes the next `-vN`, and refuses to build if a verbatim isn't in a source file or a count no longer matches the data.
- **The deck's words live in the builder** (`deck/build_case1_deck.py`, `deck/big_idea_spec.json`) as well as in the .md files. Change both.

### Case 2: Sarsabz audit + "10 feesad" → `workspace/out/Sarsabz_case2_deck-v4.pptx` (36 slides)

| Slides | Section | Source of truth |
|---|---|---|
| 1 | Cover | `comments/build_consumer_section.py` |
| 2-6 | What farmers say (1,433 comments) | `workspace/case2_audit/comments/analysis.md` |
| 7-16 | How the category talks (Sarsabz vs Engro, 200 coded uploads) | `scan/scan_findings.md`, `scan/deck/reads.yaml` |
| 17-21 | Brand Meaning Ladder | `scan/ladder_spec.json`, `scan/ladder_levels.csv` |
| 22-28 | The Big Idea: *Sarsabz proves itself on a field like yours* | `big_idea.md`, `big_idea_spec.json` |
| 29-36 | The campaign: **Dus Feesad Aur. Khet Gawah Hai.** | `concept/concept.yaml`, `concept/the_concept.md` |

Paths above are under `workspace/case2_audit/`. The full argument in prose is `workspace/case2_audit/the_audit.md`.

### Next, in priority order

1. **Ask the client.** None of these are in the brief.
   - **Both cases:** pitch date, budget, and whether Case 1 and Case 2 share a media plan.
   - **Case 1:**
     - Real TikTok results for Salam Kissan, 2021-2025.
     - Brand-health data on who people credit for Kissan Day. This is what would take the Case 1 brand insight off "provisional".
     - The rules of the 2023 YouTube giveaway, and whether the Ducky Bhai and Shehr Main Dehat vlog mentions were paid.
     - Any government notification recognising 18 December.
   - **Case 2:**
     - The demo-plot dataset behind the 10% claim.
     - Ki Jeet winners' input records.
     - Which YouTube channels `COMMENTS.xlsx` came from.
     - Competitor names and method for the brand-health study (p29).
   - **Full list:** `uv run sarsabz-pitch info`, plus `evidence.md` §7.
2. **Case 1 checks before the line is presented:**
   - Test both halves of the line with farmers. The rejected alternative ("Wa Alaikum Salam") and the reason are on deck slide 16.
   - Whether "Salam Kissan" is a registered trademark.
   - Watch Syngenta's #MaiKisanHun films, to see how far they already give the farmer a voice. Not watched yet; slide 9 says so.
   - Watch the 2025 "Salam Kissan 2025" short. Not watched yet.
3. **Case 1 evidence gaps:**
   - TikTok and Facebook Kissan Day activity. Neither can be read without logging in.
   - A Pakistani statistic, from a source we can read, on how young people see farming.
   - Whether the Express News, BOL, GNN, SAMAA, Aaj and Neo programmes name Sarsabz. Only their titles were checked.
4. **Bubber Sher.** The brief expects a "strategic vision for Sarsabz and Baber Sher" (p14), and neither case covers it.
5. **Client-facing versions of both decks.** Turn the volume down: no file names or source notes, and neutral wording for hostile verbatims. The `big-idea-slides` skill describes how.
6. **Case 2 evidence gaps** (audit §6):
   - FFC and Engro paid ads (Meta Ad Library).
   - Sarsabz and Engro content on Facebook and TikTok.
   - Comments on Sarsabz's 10% and Ki Jeet films.
   - Fatima annual report 2025.
   - NFDC Fertilizer Review 2024-25.
7. **Agency sections of the 90-minute pitch:** intro, team, awards, case studies (brief p11-12). These are agency facts, not ours to invent.
8. **Housekeeping decisions for the user:**
   - Sync the patched `build_deck.py` to `~/.claude/skills`? (See AGENTS.md.)
   - Turn the consumer, concept and Case 1 deck builders into skills?
   - Push the branch or not? Both remotes are public.

### Other deliverables in `workspace/out/` (latest versions)

| File | What |
|---|---|
| `Sarsabz_case1_big_idea_2026-09-15.docx` | Case 1 Big Idea sheet, from `case1_salam_kissan/big_idea.md` |
| `Sarsabz_case1_campaign_concept_2026-09-16.docx` | Case 1 campaign concept (line "Salam Kissan. Salam Pakistan."), from `case1_salam_kissan/the_concept.md`. The 15 Sep file has the earlier line. |
| `Sarsabz_TG_farmer_card-v2.pptx` | The Sarsabz farmer TG profile card, from `workspace/tg_profile/farmer_card.json` |
| `Sarsabz_case2_brand_audit_draft_2026-09-15-v5.docx` | The Case 2 audit, rendered from `the_audit.md` |
| `Sarsabz_farmer_conversation_analysis_2026-09-15.docx` | The Case 2 comment analysis |
| `Sarsabz_big_idea_sheet_2026-09-15-v2.docx` | Case 2 one-page Big Idea |
| `Sarsabz_campaign_concept_2026-09-15.docx` | Case 2 campaign rationale |
| `Sarsabz_brand_ladder_section-v2.pptx`, `Sarsabz_big_idea_section-v2.pptx` | Case 2 standalone sections (superseded by the combined deck) |

The Case 1 Word files carry 15 Sep in their names because the renderer takes the date from the machine clock. They were made on 16 Sep.

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
