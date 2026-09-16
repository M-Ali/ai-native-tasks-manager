# HISTORY

Newest first. The history of the SLIC and PTCL work this repo was copied from is in `_archive/slic_p78118/HISTORY.md`.

> **Resuming?** Read `README.md` → "Start here". At the end of 16 Sep 2026, both cases are built as internal drafts:
> - Case 1: `workspace/out/Sarsabz_case1_deck-v4.pptx` (19 slides)
> - Case 2: `workspace/out/Sarsabz_case2_deck-v4.pptx` (36 slides)
>
> Everything is committed on `sarsabz-pitch` (`6b6eeeb`); nothing is pushed. Next up are the client questions in README "Next".
> The pitch date and budget are still unknown.

## 2026-09-16 (later) — "What We Expect from the Pitch" (p14): the point-of-view deck

Checked both case decks against p14's seven expectations: only "innovative ideas across touchpoints" was covered, with farmer
understanding partly there. **Bubber Sher had no vision anywhere** (4 Case 2 slides, all scan findings), and there was nothing on
the portfolio, the emotional/functional balance, measurement, or long-term delivery.

New deck `workspace/out/Sarsabz_point_of_view-v3.pptx` (9 slides), built by `workspace/pov/build_pov_deck.py`. Its guard: every
figure comes from a `FACTS` entry with a source attached, and the build prints how many sourced figures were used (24).

Sources the user named, downloaded to `workspace/pov/sources/`:
- **Economic Survey 2025-26 ch.2** (retrieved PDF): agriculture 23.4% of GDP, 33.1% of employment, sector growth 2.89%; nutrient
  offtake 3,795k tonnes Jul-Mar FY2026 (+11.4%), nitrogen +14.8%, phosphate -1.9% on high prices; "Rs 100 more per 50kg bag =
  Rs 20 billion more on farmers"; Punjab Kissan Card aided cotton fertilizer use; tractor availability a constraint.
- **PBS agriculture page** (retrieved): ~24% of GDP, "half of employed labour force" - undated and **disagrees with the Survey**.
  Both are shown on the slide rather than picking the flattering one.
- **PACRA** (retrieved): Fatima 24 Jul 2026 - portfolio incl. "Bubbersher Urea and Bubbersher DAP", flagship brands Sarsabz and
  Bubbersher, Multan carve-out to Pakarab from 1 Jan 2025, "oligopolistic fertilizer industry"; Fatimafert 22 Apr 2020 - Bubber
  Sher urea, Sheikhupura plant 445,500 MT; FFC 31 Jul 2026 - urea/DAP share 56%/66% in 1HCY26, industry offtake ~9.3mln MT CY25.
- **FFC's own annual report: NOT retrieved.** ffc.com.pk returns 403 to curl, WebFetch and the headless browser (Cloudflare). FFC
  figures are cited from PACRA's FFC report instead, and no unretrieved number was put on a slide.

**The find worth the room:** the brief calls Bubber Sher "a regional brand... in North Zone" (p20); Fatima's own site calls it one
of Pakistan's oldest brands, marketed with no regional limit. The deck puts the contradiction on the page as a client question,
and writes the vision for the national-heritage reading with the tactical alternative stated.

## 2026-09-16 (later) — Case 1 checked against the brief's task AND goal; deck v4 (22 slides)

The user asked whether Case 1 covered p12's four task points, then supplied the goal paragraph. Checked by extracting the deck's
own text rather than from memory. Result: the task's celebration and ownership asks were covered and "beyond a single day" was
built, but **all three goal points appeared only on slide 2, the brief restatement** - the campaign delivered none of them, and
"rural"/"urban"/"nationwide movement" were asserted rather than mechanised.

Added (concept §5b-5c, deck slides 18-20, plus an urban row in the executions table):
- **Daily life:** "What I put on your table today" - the reply cut for the city, reviving the archive thread that demonstrably
  works on urban viewers (2025 food-blogger film 4.0M; 2020 urban shorts 388K and 382K).
- **The economy:** farmers state what their field puts in; the work carries 24% of GDP and 37.4% of employment (PES 2023-24).
- **Empowering, not just celebrating:** Ki Jeet (34 of 40 districts), 500+ demo plots (p5), the app's 800,000+ downloads and
  Sarsabz Asaan Rs.500bn (p7), and the UNDP film (13.7M) join Kissan Day for the first time. This deliberately links Case 1 to
  Case 2; keeping them apart was the wrong call against the goal's wording.
- **The movement:** the day already travels without the brand (Kissan Ittehad and Kashtkar Dost, The Nation 19 Dec 2024; FAO and
  the Federal Minister, ProPakistani 8 Jan 2025; the ICT administration's own campaign). A partner pack under the Salam Kissan
  name, one shared record, and one published participation number a year.

Deck **v4, 22 slides**; v3 kept. Word: `Sarsabz_case1_campaign_concept_2026-09-16-v2.docx`.

## 2026-09-16 (later) — `tg-profile` skill, and the farmer TG card

New skill `.claude/skills/tg-profile/` (synced to `~/.claude/skills` and the shared library; added to `sync-skills.sh`;
passes skill-creator's `quick_validate.py`). Built with `skill-creator`, the persona layout from
`data/reference/insurance middle age.pptx`, and the interests export in `data/reference/Interests.csv`.

- `scripts/interest_index.py`: affinity index = segment share / all-users share x 100, from the wide two-header-row export.
  It prints index AND volume, refuses to rank interests under `--min-users` (default 50), and prints the dataset's own
  header so the card can name the population.
- `scripts/build_tg_card.py`: one-slide card. Refuses to build on an unevidenced pain point or trigger, a `%` with no
  source, an interests source naming no population, or text that would overflow its box.
- `references/card-anatomy.md` documents the blocks and what to copy (and not) from the insurance card: its media
  percentages carry no source, which is why the builder rejects that pattern.

**Decisions the user made:** the card covers the Sarsabz farmer only; `Interests.csv` (brandsynario.com's own web audience,
35,681 urban visitors, Jan-Sep 2026) is kept for a future urban card because it does not describe farmers; and the media
block shows evidenced channels with no invented percentages.

**Deliverable:** `workspace/out/Sarsabz_TG_farmer_card-v4.pptx`, spec `workspace/tg_profile/farmer_card.json`. Pain points
come from the 1,433 agronomy comments (counts verified against `theme_counts.csv`) and the 1,572 Salam Kissan comments.
v1 is kept; its render showed the media panel overflowing, fixed in v2.

**v3, on the user's instruction:** the interests block now carries six categories from a second GA4 export the user supplied,
`data/reference/Punjab.csv` (brandsynario.com visitors **in Punjab**, male 25-44: 390 of 1,472 users) - Avid Political News
Readers i264, Cricket i185, Auto i183, Business News i161, Avid Investors i159, Avid News Readers i126. The media block follows
the insurance template's frame (traditional TV/radio/OOH plus the social platforms), with no percentages, since no rural survey
exists. **The caveat stands and is printed on the card:** this is a website audience in Punjab, not a farmer sample - the
taxonomy has no agriculture category at all, and against the national file Punjab skews *more* digital (Mobile Enthusiasts
i159, Technophiles i131), which is the opposite of a rural signal. Treat the indices as directional only. (v3 clipped its media panel; the builder's fit check now counts the lines each
bullet wraps to, rather than total characters, and v4 is the clean render.)

## 2026-09-16 (later) — Case 1 line changed, deck v3

The user challenged the proposed line "Salam Kissan. Wa Alaikum Salam, Pakistan." Checking it against the corpus settled it: the
greeting appears in 3 of 1,516 comments and never as a reply to the campaign, while **"Salam Pakistan" appears in 19**, unprompted,
including under "We need to make this permanent not temporary. They are us and we are them. SALAM PAKISTAN" (87 likes).

- **Line is now "Salam Kissan. Salam Pakistan."**, with the mechanic named **Kissan ka Jawab** (the farmer's answer) and the UGC tag
  **#KissanKaJawab**. The old line stays on the line slide as considered-and-rejected, with the reasons: religious register, a farmer
  who isn't Muslim can't say it in character, the correct form is the longer "Wa Alaikum Assalam", and the corpus evidence above.
- `analyse.py` now codes `salam_pakistan`, so the 19 is reproducible; the builder checks the count before it will build.
- **Deck: `workspace/out/Sarsabz_case1_deck-v4.pptx`** (v2 kept). Word: `Sarsabz_case1_campaign_concept_2026-09-16.docx`.
- Four Big Idea options inside the farmer-voice territory were offered (the reply; a published record of farmers' asks; farmers write
  the anthem; the farmer as expert). The deck still carries the reply; the user hasn't chosen between them.

## 2026-09-16 — Case 1 built: evidence, Big Idea, concept, deck (commit `6b6eeeb`)

Under the user's rule **"no made up data, no assumptions"**. Everything is in `workspace/case1_salam_kissan/`; `evidence.md` is the summary.
- `archive/`: 76 Salam Kissan uploads coded (format, speaker, audience, window): 79% of uploads in 1 Dec-15 Jan; a farmer speaks in 9 of 76.
- `comments/`: 1,572 comments on 18 films, all read. `analyse.py` counts them with regexes plus a list of hand-picked verbatim
  fragments that must each match exactly one comment. The 2023 thread is inflated by a giveaway (31 winner replies) and vlog referrals.
  Complaints: 19 of 1,516. The skill's `load_comments.py` misread the clean CSV (6,219 "comments"), so its output is set aside in `comments/_rejected/`.
- Press checks on who owns the day, plus PES and Census facts. TikTok is unreadable logged-out, and its 1.5M UGC / 4.19B views figures are unverified and not used.
- `scan/`: YouTube search (7 queries, 254 videos) and metadata for 36 others. Syngenta runs Kisan Day films (2023-24); JPL and Rizq Foods
  use "Salam Kissan"; no FFC or Engro campaign was found; one agri channel calls the day "sirf tv show".
- **Big Idea** (`big_idea.md`): "Salam Kissan gets its reply: the farmer answers Pakistan's salute in his own voice, through his own
  family." Brand insight is provisional (no attribution data).
- **Concept** (`the_concept.md`): proposed line "Salam Kissan. Salam Pakistan." with the mechanic "Kissan ka Jawab", with executions for ATL, UGC, BTL, PR, events and on-ground. It asks Fatima to put farmers' complaints on air and to name
  the brand with the day.
- Word files: `workspace/out/Sarsabz_case1_big_idea_2026-09-15.docx`, `Sarsabz_case1_campaign_concept_2026-09-15.docx`.
- **Deck:** `workspace/out/Sarsabz_case1_deck-v4.pptx`, 19 slides: cover; brief; archive; comments ×2; category scan; Big Idea ×7
  (big-idea-slides skill); campaign ×6. Built by `workspace/case1_salam_kissan/deck/build_case1_deck.py`, which refuses to build if
  any verbatim isn't in a source file or a count no longer matches the data. v1 is kept: its render check showed line overlap and
  table overflow, fixed in v2.
- **Next:** client questions (TikTok data, attribution, giveaway rules, budget, pitch date); a client-facing version if wanted.

## 2026-09-16 — committed to git

The whole project folder is committed on branch **`sarsabz-pitch`** (repo root `D:\Personal`, left off `master`). Only `.venv`, caches
and `~$` Office lock files are excluded. Nothing is pushed. **Both remotes (`ai-native-tasks-manager`, `-new`) are PUBLIC**, and this
branch holds the client's brief, the farmer comments and archived SLIC/PTCL work, so don't push it without deciding that first.

- `86cfdf3`: 2,189 files. The inherited SLIC `.gitignore` rule `workspace/out/` silently left out every deck and Word file.
- Second commit: that rule removed, so all deliverables in `workspace/out/` (and `_archive/slic_p78118/workspace/out/`) are tracked.

The 321 uncommitted changes in other projects under `D:\Personal` were deliberately not touched.

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
