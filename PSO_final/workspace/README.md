# PSO Lubricants Deep-Dive Reports

Standalone report-generation scripts that sit on top of the core pipeline
(`src/pso/` — see `../AGENTS.md`). They reuse `pso.ingest.load()` for data
but are **not** part of the `pso run` CLI; they're run directly.

## Quick start

```bash
# Regenerate everything from the current data file
uv run python workspace/run_all_reports.py

# Regenerate everything from a NEW period's file (e.g. FY27)
uv run python workspace/run_all_reports.py --input "data/input/Working File Retail Fuels Data FY27.xlsx"
```

That single command runs all 8 scripts below in sequence and writes every
output into `reports/`, automatically tagged with the period read from the
source file (e.g. `_10M_FY26`, `_10M_FY27`) — so a new run never overwrites
an older period's reports.

## What gets generated

| Script | Output | Contents |
|---|---|---|
| `lubes_report.py` | `PSO_Lubricants_Report_<period>.docx` | Full lubricants business report — revenue, volume, margin, regional, city, station analysis |
| `lubes_stations_analysis.py` | `PSO_Lubes_Stations_Analysis_<period>.png` | Standalone chart — stations & volume by category and top-15 cities |
| `lubes_stations_report.py` | `PSO_Lubes_Stations_Report_<period>.docx` | Word version of the above with embedded charts and tables |
| `city_profiles.py` | `city_profiles/PSO_<City>_Lubes_Profile_<period>.pptx` (10 cities) | 5-slide **revenue**-based city profile decks |
| `city_profiles_volume.py` | `city_profiles_volume/PSO_<City>_Lubes_Vol_Profile_<period>.pptx` (11 cities) | 5-slide **volume**-based city profile decks |
| `lubes_vol_table.py` | `PSO_Lubes_Vol_Top20_Cities_Table_<period>.docx` | Landscape table — top 20 cities, volume + category breakdown |
| `lubes_vol_uplift.py` | `PSO_Lubes_Vol_Uplift_Table_<period>.docx` + `PSO_Lubes_Uplift_Scenarios_<period>.pptx` | Volume uplift potential from the "Where to Focus" initiatives — conservative & optimal scenarios |
| `national_vol_slide.py` | `PSO_National_Lubes_Vol_Station_Profile_<period>.pptx` | Single slide — national lubricants volume station performance (all retail stations aggregated) |

All city selections (top 10, top 15, top 20) and every figure are computed
live from the data each run — nothing about *which* cities or *what*
numbers appear is hardcoded. Only formatting and scenario assumptions
(colors, the 25%/55% conservative/optimal multipliers in
`lubes_vol_uplift.py`) are fixed in code — see comments at the top of each
script.

## YoY comparison methodology — SPLY vs LY 12M

All **YoY % change figures** in every report are calculated against **Same Period Last Year (SPLY)** — a like-for-like comparison of the same number of months (e.g. 10M CY vs 10M SPLY). This avoids the apples-to-oranges distortion of comparing 10 months of CY data against a full 12-month LY.

**How SPLY is derived:** The source Excel file contains `%SPLY*` columns (`%SPLYSalesLtr`, `%SPLYGRS`, `%SPLYNetMargin`, `%SPLYDisc`, `%SPLYPMargin`) — pre-calculated percentage changes vs the same period last year. At load time, `pso.ingest` derives absolute SPLY values as:

```
SPLY = CY / (1 + %SPLY / 100)
```

These become the `SalesLtr_SPLY`, `SalesGRS_SPLY`, `NetMargin_SPLY` etc. columns used by all downstream scripts.

**LY 12M values are still shown in tables**, clearly labelled `LY 12M`, so the full prior-year scale is visible for context. But the `% change` column always uses the SPLY denominator, not LY 12M.

New stations with no prior-period data have `%SPLY = NaN` → their SPLY is set to 0, so their % change shows as `+∞` or `N/A` depending on the context.

## How period auto-tagging works (`_pso_common.py`)

- `INPUT_PATH` — reads the `PSO_INPUT` env var; falls back to
  `data/input/Working File Retail Fuels Data.xlsx` if unset. The
  orchestrator sets this for you when you pass `--input`.
- `get_period_label(df)` — builds the human-readable label shown on report
  covers/footers (e.g. `"10M FY26 (generated June 2026)"`) from the source
  file's Excel **sheet name** — the same value `pso.ingest.load()` reports
  as `Period:` when it loads.
- `out_path(stem, ext, df)` — appends `_<period>` to every output filename.

You generally don't need to touch this file. It's only consulted when you
run scripts with a different input file.

## Running a single script

Each script can still be run on its own:

```bash
uv run python workspace/lubes_report.py
PSO_INPUT="data/input/Working File Retail Fuels Data FY27.xlsx" uv run python workspace/city_profiles.py
```

## Onboarding a new period (e.g. FY27) — checklist

1. **Check schema/quality first** — run the core pipeline's quality check
   against the new file before trusting any deep-dive numbers:
   ```bash
   uv run python -m pso.main quality --input "data/input/<new file>.xlsx"
   ```
   If it flags new unmapped city names or product codes, add them to
   `CITY_NORM` / `PRODUCT_SEGMENTS` in `src/pso/config.py` (see `AGENTS.md`).
   Otherwise those rows fall back to their raw city name (or `Unknown`
   segment) and won't roll up correctly into city/category totals.
2. **Run everything**:
   ```bash
   uv run python workspace/run_all_reports.py --input "data/input/<new file>.xlsx"
   ```
3. New, period-tagged files land in `reports/` (and its `city_profiles*`
   subfolders) alongside whatever was there before. Nothing is deleted or
   overwritten.
4. **City list may shift.** Profiles are generated for whichever cities
   rank in the new top-10/11 by volume — if a new city enters the top
   ranks, it gets a fresh profile; a city that drops out simply won't be
   regenerated (its old deck stays in the folder, now stale).

## Cleanup — files safe to delete (optional, not required)

The refactor above was tested against the existing FY26 file, which left
old **untagged** files sitting next to the new **tagged** ones — pure
duplicates from before this change. Safe to delete if you want a tidier
folder; nothing depends on them:

- `reports/PSO_Lubricants_Report.docx`
- `reports/PSO_Lubes_Stations_Analysis.png`
- `reports/PSO_Lubes_Stations_Report.docx`
- `reports/PSO_Lubes_Vol_Top20_Cities_Table.docx`
- `reports/PSO_Lubes_Vol_Uplift_Table.docx`
- `reports/PSO_Lubes_Uplift_Scenarios.pptx`
- `reports/city_profiles/PSO_<City>_Lubes_Profile.pptx` (10 files, untagged)
- `reports/city_profiles_volume/PSO_<City>_Lubes_Vol_Profile.pptx` (11 files, untagged)

Files like `Lubes_report.pptx`, `Master.pptx`, `PSO_Karachi_District_Volume.pptx`,
`PSO_Product_Revenue.pptx`, `PSO_Top5_Cities_Volume.pptx`, `PSO_Report_10M_FY26.*`
are **not** part of this list — they come from other one-off scripts
(`city_slides.py`, `product_revenue_slide.py`, etc.) untouched by this
refactor and have no tagged counterpart.

### Does keeping the untagged duplicates hamper anything?

No. No script reads its own previous output — every run starts fresh from
the source Excel file. The untagged files are dead weight, not a
dependency. The only downside is clutter: at a glance it's not obvious
which `PSO_Lubricants_Report*.docx` is current (answer: the one with a
period suffix is always the latest for that period).

### Backup status

- **`reports/`** is in `.gitignore` — it's treated as disposable, regenerable
  output, not source-controlled. That's normal; nothing here needs a backup
  because rerunning the scripts reproduces it exactly.
- **`workspace/*.py`** (the actual report logic) is **git-tracked** — the
  entire `workspace/` folder was committed to the repo (`543f2b4`). All
  subsequent edits (SPLY implementation, legend fixes, national slide) are
  also committed. Safe to fall back to any prior version via `git log
  workspace/`.
