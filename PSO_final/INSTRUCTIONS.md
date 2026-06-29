# PSO Analytics — How to Run

## The Two Apps at a Glance

There are **two separate, independent apps** in this project. They do not call each other.

| | App 1 — Core Pipeline | App 2 — Report Scripts |
|---|---|---|
| **Entry point** | `pso.main` | `workspace/run_all_reports.py` |
| **What it does** | Ingest → Analyse → AI Narrative → Excel | Ingest → Word & PowerPoint reports |
| **Needs API key?** | Yes (for AI step; can skip) | No |
| **Output** | Multi-sheet Excel workbook | Word docs, PowerPoint decks, PNGs |
| **Output folder** | `reports/YYYY-MM-DD/` | `reports/` and subfolders |
| **Run time** | ~2–3 min (with AI) / ~30 sec (no AI) | ~1–2 min |

They share the same source Excel file and both run `pso.ingest.load()` internally — but beyond that they are completely independent. Running one does **not** run the other.

---

## Prerequisites

- **uv** must be installed (package/environment manager). Run `uv --version` to confirm.
- All commands must be run from the **project root**: `D:\Personal\PSO_final\`
- API key is only needed for App 1 with AI. If unset, use `--no-ai` to skip that step.

---

## App 1 — Core Pipeline (`pso.main`)

Runs the 5-agent pipeline: **Ingest → Analysis → Lubes Analysis → AI Narrative → Excel Report**.

```bash
# Standard run (with AI narrative — needs Anthropic API key)
uv run python -m pso.main run

# Skip AI (faster, no API key needed — still produces full Excel analysis)
uv run python -m pso.main run --no-ai

# Specify a different input file
uv run python -m pso.main run --input "data/input/Working File Retail Fuels Data FY27.xlsx" --no-ai

# Data quality check only (fast — just flags issues in the source file)
uv run python -m pso.main quality
uv run python -m pso.main quality --input "data/input/<new file>.xlsx"
```

**Output:** A dated Excel workbook in `reports/YYYY-MM-DD/` with 14 sheets covering
portfolio summary, diesel, petrol, lubricants, city pareto, regional gaps, and AI observations.

---

## App 2 — Report Scripts (`workspace/run_all_reports.py`)

Runs all 8 report-generation scripts in sequence, producing Word docs and PowerPoint decks.

```bash
# Standard run (uses default input file)
uv run python workspace/run_all_reports.py

# Specify a different input file (e.g. for a new period)
uv run python workspace/run_all_reports.py --input "data/input/Working File Retail Fuels Data FY27.xlsx"
```

**Output files produced (all in `reports/`):**

| File | Contents |
|---|---|
| `PSO_Lubricants_Report_<period>.docx` | Full lubricants business report |
| `PSO_Lubes_Stations_Analysis_<period>.png` | Stations & volume chart |
| `PSO_Lubes_Stations_Report_<period>.docx` | Word version of above |
| `PSO_Lubes_Vol_Top20_Cities_Table_<period>.docx` | Top 20 cities volume table |
| `PSO_Lubes_Vol_Uplift_Table_<period>.docx` | Where-to-focus volume uplift table |
| `PSO_Lubes_Uplift_Scenarios_<period>.pptx` | Conservative & optimal uplift scenarios slide |
| `PSO_National_Lubes_Vol_Station_Profile_<period>.pptx` | National station performance slide |
| `city_profiles/PSO_<City>_Lubes_Profile_<period>.pptx` | Revenue-based city profile (10 cities) |
| `city_profiles_volume/PSO_<City>_Lubes_Vol_Profile_<period>.pptx` | Volume-based city profile (11 cities) |

Output filenames are **auto-tagged with the period** read from the source file's sheet name
(e.g. `_10M_FY26`, `_10M_FY27`). A new period's run never overwrites a prior period's reports.

---

## Running a Single Script

Each workspace script can be run on its own:

```bash
uv run python workspace/lubes_report.py
uv run python workspace/national_vol_slide.py
uv run python workspace/city_profiles_volume.py
```

To point a single script at a different input file:

```bash
# Windows PowerShell
$env:PSO_INPUT="data/input/Working File Retail Fuels Data FY27.xlsx"; uv run python workspace/lubes_report.py
```

---

## New Period Workflow (e.g. FY27)

1. **Drop the new Excel file** into `data/input/`

2. **Check data quality first** — flags unmapped city names or new product codes:
   ```bash
   uv run python -m pso.main quality --input "data/input/<new file>.xlsx"
   ```
   If new city names are flagged, add them to `CITY_NORM` in `src/pso/config.py`.

3. **Run the core pipeline** (optional — for Excel report + AI narrative):
   ```bash
   uv run python -m pso.main run --input "data/input/<new file>.xlsx" --no-ai
   ```

4. **Run all Word/PPTX reports:**
   ```bash
   uv run python workspace/run_all_reports.py --input "data/input/<new file>.xlsx"
   ```

5. New files land in `reports/` tagged with the new period. Old period files are untouched.

---

## Quick Reference

```bash
# Full Word/PPTX report refresh (most common command)
uv run python workspace/run_all_reports.py

# Full Excel + AI refresh
uv run python -m pso.main run --no-ai

# New period — replace <new file> with actual filename
uv run python workspace/run_all_reports.py --input "data/input/<new file>.xlsx"

# Data quality check on new file
uv run python -m pso.main quality --input "data/input/<new file>.xlsx"
```

---

## Key Things to Remember

- **Two apps, run separately.** `run_all_reports.py` does not trigger `pso.main` and vice versa.
- **No manual edits needed between periods.** Period label and output filenames are auto-detected from the source file's sheet name.
- **`reports/` is not backed up by git** — it is regenerable output. The scripts in `workspace/` and `src/` are git-tracked.
- **Old period files are never deleted.** Each run adds new tagged files alongside existing ones.
- **City profiles regenerate for whoever ranks in the new top cities.** If a city drops out of the top ranks in FY27, its FY26 deck stays in the folder but won't be refreshed.
