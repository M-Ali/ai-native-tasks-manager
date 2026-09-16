# AGENTS.md — working conventions for this repo

## What this project is

`sarsabz-pitch` supports one live pitch: **Fatima Fertilizer's appointment of a strategic and creative agency**
for Fatima Fertilizer, Pak Arab and FatimaFert (brands: Sarsabz, Bubber Sher, Pakarab CAN). The source of truth
is `data/Creative Pitch Deck (1).pdf`.

- 90-minute presentation: agency, team, awards, case studies, brand understanding, Case 1 and Case 2 solutions.
- **Case 1**: Salam Kissan 2026. *Evidence, Big Idea and concept drafted* (`workspace/case1_salam_kissan/`); deck `workspace/out/Sarsabz_case1_deck-v3.pptx` built by `deck/build_case1_deck.py` (verbatim and count guards). No made-up data, no assumptions. In the deck builder, every new verbatim goes through `V()`, which checks it against the sources, and every translation through `T()`, which labels it.
- **Case 2**: Sarsabz brand audit and a tactical campaign for "10 feesad se bhi ziada izafi paidawar". *Built:* the current deck is
  `workspace/out/Sarsabz_case2_deck-v4.pptx`. See README "Start here".
- **No pitch date, no budget** in the brief. Do not invent either; they are open questions in `brief.py`.

This folder was copied from earlier pitches (SLIC life insurance, PTCL telco). That work is in `_archive/` and must
not leak into this one: no insurance or telco vocabulary, brands, examples or figures in live files.

## Ground rules

1. **The brief is the authority.** Every item in `brief.py` carries the brief page it was read from. Re-read the
   page before changing an item. A requirement the brief doesn't state goes in `OPEN_QUESTIONS`.
2. **Look at the image-only pages.** Pages 15-26 and the brand-health charts (28-31) have no text layer.
   `sarsabz-pitch parse` renders them.
3. **Never overwrite a deliverable.** Every builder writes the next free `-vN`, except `build_deck.py` given an explicit
   `.pptx` path, which overwrites. Point that one at the scratchpad and let a later builder write into `workspace/out/`.
4. **`_archive/` is read-only.** Move things in; never edit or delete inside it.
5. **`data/` belongs to the user.** Open a file and check its modified date before moving it. A filename matching an old
   client's file is not the same file: `COMMENTS.xlsx` was misfiled this way on 15 Sep 2026. When in doubt, ask.
6. **Check a skill doesn't already exist before creating one.** Search `~/.claude/skills/`,
   `D:/Personal/Skills_aug2026/skills-main/skills/` and `D:/Personal/Strategy_PTCL_Sep/.claude/skills/`. On 15 Sep,
   `brand-laddering`, `big-idea` and `big-idea-slides` all already existed when new ones were requested.

## Skills

| Skill | Lives in | Used for |
|---|---|---|
| `brief-to-plan` | project + `~/.claude` | Brief → plan (not yet run on this brief; needs the pitch date) |
| `consumer-brand-review` | `~/.claude` | Category research method |
| `category-creative-scan` | project (patched) + `~/.claude` | Presence sweep, coding schema, `build_deck.py` for the scan section |
| `competitor-comms-audit` | project + `~/.claude` | Territory method |
| `tg-profile` | project + `~/.claude` + library | TG profile card: `interest_index.py` for affinity indices, `build_tg_card.py` for the slide |
| `comment-analysis` | project | Loader, lexicon coder, the farmer analysis |
| `brand-laddering` | `~/.claude` + shared library only | `compute_levels.py`, `build_ladder_section.py` |
| `big-idea` | copied from PTCL to `~/.claude` + shared library | The three-insight method |
| `big-idea-slides` | `~/.claude` + shared library + PTCL | `build_big_idea_section.py` |
| `campaign-concept` | project + `~/.claude` | The chain and six tests (no slide builder exists) |

- **`sync-skills.sh` does `rm -rf` and then copies** each listed skill from the project to `~/.claude/skills/` and the
  library. Never add a skill that exists only in `~/.claude` to its list; that would delete it.
- **The project copy of `category-creative-scan/scripts/build_deck.py` is patched and not synced.** It swaps the hardcoded insurance stat
  for the objection rate, and adds a per-scan `_method` override. The original is in `_archive/draft_skills/`. Syncing would push
  the patch to other pitch folders; that is the user's call.
- **Project-level builders (no skill yet):** `workspace/case2_audit/comments/build_consumer_section.py` and
  `workspace/case2_audit/concept/build_concept_section.py`. Both are candidates to become skills.

## Build and check

The deck is built, never hand-edited. The full chain and commands are in README "Rebuild the deck". Order:
scan (`build_deck.py`) → consumer and cover (`--base --front`) → ladder (`--base`) → Big Idea (`--base`) → concept (`--base`).

- **Every builder has guards. Fix the spec, never the guard.** The ladder builder refuses brand labels that disagree with
  computed levels. The Big Idea builder refuses an idea without alternatives. The consumer builder refuses a quote not
  verbatim in `comments.csv`. The concept builder refuses archive titles or brief phrases it can't find.
- **Render every changed slide and look at it.** python-pptx does not show overflow. PowerPoint's early-bound COM type
  library is broken on this machine, so use late binding in PowerShell (`New-Object -ComObject PowerPoint.Application`,
  then `InvokeMember` for Presentations → Open → `SaveAs(dir, 18)`, which writes one PNG per slide). Render checks on 15 Sep
  caught an overlapping headline and a method slide describing the wrong study.
- **Spec copies for the combined build** go to the scratchpad; the original specs (`scan/ladder_spec.json`,
  `big_idea_spec.json`) keep their section-deck output paths.

## Category facts that change the work

Sourced in `workspace/case2_audit/the_audit.md`. Keep the precision; each was an error once.

- Fatima Group makes **the only CAN and 93% of NP**. Not "the only NP": Engro holds 6% of NP (CCP Table 11, DRAFT, NFDC 2022-23).
- FFC holds 44% of urea and 61% of DAP but has no official YouTube channel. Engro is the only competitor with owned video to code.
- The brief's "exceeded competition's TOM awareness by 6%" (p29) holds against Competitor 2 only; Competitor 1 is at 62%.
- Ki Jeet films claim the competition "proved" 10%. It shows high yields, not a 10% difference. Dawn (3 Aug 2025) contradicts
  "all winners used Sarsabz" for Chakwal.
- NP retails well below DAP and CAN below urea (BMA, reported). Sarsabz makes no price claim in its 24-month window.

## Evidence rules

7. **Code from the creative, not the caption.** Build contact sheets and look at every item before coding.
8. **`unclear` is a valid code; a guessed code is not.**
9. **Views on a brand channel are bought media weight.** Use them to show where the brand spends, never as organic
   interest or share of voice.
10. **Never present share of voice from a capped or equal-N sample.**
11. **Retrieved vs reported.** Mark which. Dawn "advertorial" pieces are Fatima's own words.
12. **Search for handles; never guess them.** A "Bubber Sher" Facebook page found this way was a rock band.
13. **Absence is a finding.** Record it.
14. **Space out YouTube requests.** A bot check appeared after ~420 requests in one session.
15. **Verify every quote and figure by script before it goes on a slide**, and label translations.

## Writing sections

16. **Markdown is the source; Word is a render** (`sarsabz-pitch doc`). Never edit the `.docx`.
17. **Say what the evidence can't show**, and distinguish a data request from an assumption.
18. **Check claims a rival could challenge.** A checkable overclaim costs more than a smaller true one.

## Change log

Record every material change in `HISTORY.md`, newest first, dated.
