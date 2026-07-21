# PSO OMC Analytics — Agent Roles & Responsibilities

## Purpose
This file defines the agent roles within the PSO analytics pipeline. Each agent has a
single responsibility. No agent invents data or assumptions — every output is traceable
to a row in the source file.

---

## Agent 1 — Ingest Agent (`src/pso/ingest.py`)

**Responsibility:** Load raw Excel, validate schema, normalize city names, classify
products into segments, and emit a clean DataFrame that all downstream agents consume.

**Inputs:** Raw `.xlsx` file path
**Outputs:** `pd.DataFrame` — clean, validated, column-enriched

**Rules:**
- Fail loudly if expected columns are missing — never silently skip.
- City normalization uses a deterministic lookup table. If a city is not in the table it
  is kept as-is and flagged in the data quality report. No fuzzy guessing.
- Product segment assignment (`Diesel / Petrol / Lubricants / Other Fuels / LPG /
  Aviation / LNG / Chemicals`) is driven by `PRODUCT_SEGMENTS` in `config.py`.
  If a product code is not mapped it raises a `ValueError`.
- Lube sub-category assignment uses the `Category` column. Blank `Category` for non-lube
  products is expected and not flagged as an error.
- Adds derived columns: `FuelSegment`, `LubeCategory`, `CityNorm`, `IsRetail`,
  `IsInternational`.

**Data quality checks emitted:**
- Rows with null `Sales _Org_ Desc.2`
- Rows with null `Sales office Region` for Retail rows
- City names not in normalization table (counts only — not blocking)
- Negative CY values (data anomaly flag)

---

## Agent 2 — Analysis Agent (`src/pso/analyze.py`)

**Responsibility:** Produce all aggregated tables from the clean DataFrame. Pure pandas —
no AI, no narrative, no opinions. Returns a dict of named DataFrames.

**Inputs:** Clean `pd.DataFrame` from Ingest Agent
**Outputs:** `dict[str, pd.DataFrame]` — one key per analysis table

**Analysis tables produced:**

| Key | Description |
|-----|-------------|
| `portfolio_summary` | All Sales Org channels: GRS, Volume, Net Margin CY vs SPLY |
| `retail_segment_split` | Retail: Fuels vs Lubricants totals and ratios |
| `diesel_by_region` | HSD + LDO: Volume, GRS, Net Margin by Region |
| `diesel_by_city` | HSD + LDO: same, by City (top 50 by volume) |
| `diesel_pareto` | Cumulative volume share by city — marks 50% / 80% cutoffs |
| `petrol_by_region` | PMG + R95: Volume, GRS, Net Margin by Region |
| `petrol_by_city` | PMG + R95: same, by City (top 50 by volume) |
| `petrol_pareto` | Cumulative volume share by city — marks 50% / 80% cutoffs |
| `lubes_overview` | All Lube categories combined: CY vs SPLY trend |
| `lubes_by_category` | Each lube type (DEO/PCMO/MCO/LOW GRADE/Greases/etc.) |
| `lubes_by_region` | Lube categories × Region matrix |
| `lubes_by_city` | Lube categories × City (top 30 by volume) |
| `lubes_discount_erosion` | Discount/Ltr vs Primary Margin/Ltr vs Net Margin/Ltr per category |
| `lubes_customer_segments` | Top corporate groups buying lubes — volume and margin |
| `region_performance` | Region × Segment matrix: volume share, margin/ltr, YoY change |
| `city_pareto_all` | All Retail: cumulative contribution — value and volume |
| `underperforming_regions` | Regions where volume declined AND margin compressed |
| `data_quality_report` | Row counts for each data issue flagged by Ingest Agent |

**Rules:**
- All percentage changes: `(CY - SPLY) / abs(SPLY)` — handle zero SPLY gracefully.
- Never drop rows silently. If a row cannot be classified it goes to `_unclassified`.
- Volume = `SalesLtr_CY/SPLY`. Margin = `NetMargin_CY/SPLY`. Revenue = `SalesGRS_CY/SPLY`.
- Pareto cutoffs are computed from data — no hardcoded thresholds.

---

## Agent 3 — Lubes Agent (`src/pso/lubes_analyze.py`)

**Responsibility:** Dedicated lubricants analysis. Goes deeper than the Analysis Agent on
PSO's core problem area.

**Inputs:** Clean `pd.DataFrame` filtered to Retail Business + Lubricant products
**Outputs:** `dict[str, pd.DataFrame]` — lube-specific tables

**Analysis tables produced:**

| Key | Description |
|-----|-------------|
| `lube_category_trend` | Each category CY vs SPLY: volume, GRS, primary margin, discount, net margin |
| `lube_margin_decomp` | Primary Margin → Discount → Net Margin waterfall per category |
| `lube_region_category` | Pivot: Region × Lube Category (volume) |
| `lube_city_category` | Top 20 cities × Lube Category (volume) |
| `lube_top_customers` | Top 50 customers by lube volume CY, with category mix |
| `lube_declining_markets` | Cities where lube volume declined YoY — sorted by absolute loss |
| `lube_growing_markets` | Cities where lube volume grew YoY — sorted by absolute gain |
| `lube_discount_city` | Average Discount/Ltr by city — identifies where most is given away |
| `lube_margin_per_ltr` | NetMargin/Ltr by category × region — identifies value pockets |
| `low_grade_analysis` | Deep dive on LOW GRADE specifically — who buys, where, margin |

**Rules:**
- Separate analysis for each lube `Category` value — never aggregate across categories
  unless explicitly producing an "all lubes" summary.
- Flag any category where Net Margin/Ltr CY < Net Margin/Ltr SPLY (margin compression).
- Flag any category where Discount/Ltr CY > Discount/Ltr SPLY (increased discounting).

---

## Agent 3b — Premium Fuel Agent (`src/pso/premium_fuel_analyze.py`)

**Responsibility:** R95 (premium) vs PMG (regular) petrol deep-dive — the same depth of
treatment the Lubes Agent gives lubricants, applied to PSO's existing premium fuel
product. Rather than speculating about a hypothetical new high-end fuel launch, this
agent uses R95's real, already-selling footprint as the proof-of-concept and answers
"where to launch/expand" directly from station-level data.

**Inputs:** Clean `pd.DataFrame` filtered to Retail Business + Petrol products (PMG, R95)
**Outputs:** `dict[str, pd.DataFrame]` — premium-fuel-specific tables

**Analysis tables produced:**

| Key | Description |
|-----|-------------|
| `premium_product_trend` | PMG vs R95 national: volume, GRS, margin, CY vs SPLY |
| `premium_margin_decomp` | Primary → Discount → Net Margin per litre, PMG vs R95, plus the R95-minus-PMG uplift row |
| `premium_by_region` | R95 share of regional petrol volume, ranked |
| `premium_by_city` | R95 penetration % by city, ranked — proof-of-demand markets |
| `premium_station_mix` | Station-level R95 vs PMG volume + R95 % of the station's own petrol volume |
| `premium_whitespace_stations` | Active PMG stations with zero R95 — ranked by PMG volume (launch priority) |
| `premium_growing_markets` / `premium_declining_markets` | Cities where R95 volume grew/declined vs SPLY |
| `premium_customer_segments` | Corporate-group / customer mix buying R95 |

**Rules:**
- CY vs SPLY throughout — same basis as every other agent (see invariant #4 below).
- Whitespace = active PMG volume with zero R95 — this is the primary "where to launch"
  output; do not confuse it with `premium_by_city`/`premium_by_region`, which show where
  R95 already sells (proof of demand, not launch candidates).
- No external market data (competitor fuels, regulation) — internal transaction data only.

---

## Agent 4 — AI Narrator Agent (`src/pso/ai_insights.py`)

**Responsibility:** Takes structured analysis tables (not raw data) and calls the Claude
API to produce business narrative, identify patterns, and recommend actions. All Claude
prompts include the actual numbers — no vague questions.

**Inputs:** `dict[str, pd.DataFrame]` from Analysis Agent + Lubes Agent
**Outputs:** `dict[str, str]` — one narrative block per analysis section

**Sections narrated:**
1. `exec_summary` — Portfolio-level: what's happening across all channels
2. `retail_fuels` — Diesel and Petrol performance, regional differences, city concentration
3. `lubes_problem` — Root cause analysis of lubricants underperformance from data
4. `regional_gaps` — Which regions underperform, what the data shows as likely drivers
5. `recommendations` — Prioritized action list, each tied to a specific data finding

**Rules:**
- Claude receives DataFrames serialized as markdown tables — not raw CSV rows.
- Each prompt explicitly states: "Base your analysis only on the numbers provided."
- Responses are structured (numbered findings, then recommendations).
- If Claude cannot identify a pattern from the numbers alone it says so — no hallucination.
- Token budget per section: ~1,500 tokens output. Concise, not exhaustive.

---

## Agent 5 — Report Agent (`src/pso/excel_report.py`)

**Responsibility:** Assemble all analysis tables and AI narratives into a formatted,
multi-sheet Excel workbook. One workbook per run, versioned by date.

**Inputs:**
- `dict[str, pd.DataFrame]` from Analysis Agent + Lubes Agent + Premium Fuel Agent
- `dict[str, str]` from AI Narrator Agent

**Outputs:** `.xlsx` file in `reports/YYYY-MM-DD/`

**Excel sheet structure:**

| Sheet | Contents |
|-------|----------|
| `00_Summary` | Portfolio scorecard — all channels, KPIs, traffic-light formatting |
| `01_Retail_Split` | Fuels vs Lubes side by side |
| `02_Diesel` | HSD+LDO: region table, city top-50, Pareto chart data |
| `03_Petrol` | PMG+R95: region table, city top-50, Pareto chart data |
| `04_Lubes_Overview` | All lube categories: CY vs SPLY, margin decomposition |
| `05_Lubes_DEO` | DEO: region × city breakdown |
| `06_Lubes_PCMO` | PCMO: region × city breakdown |
| `07_Lubes_MCO` | MCO: region × city breakdown |
| `08_Lubes_LowGrade` | LOW GRADE: region × city breakdown |
| `09_Lubes_Other` | Greases, Industrial Grade, Others |
| `10_City_Pareto` | All products: 50%/80% concentration cutoffs |
| `11_Regional_Gaps` | Underperforming regions with data evidence |
| `12_AI_Observations` | Claude narrative per section + prioritized recommendations |
| `13_City_Opportunity` | Lube mix + fuel discount opportunity sizing by city |
| `14_Premium_Fuel_R95` | R95 vs PMG: national trend, margin decomp, region/city penetration, growers/decliners |
| `15_Premium_Fuel_Stations` | Station-level R95 mix + whitespace (active PMG stations with zero R95) |
| `16_Data_Quality` | Flagged issues: nulls, unnormalized cities, anomalies |

**Formatting rules:**
- Green fill: YoY growth > 5%
- Red fill: YoY decline > 5%
- Yellow fill: YoY change within ±5%
- All currency in PKR billions (B) or millions (M) — no raw numbers.
- Volume in Million Liters (M Ltrs) — not raw liters.
- Freeze top row on every sheet.
- Tables use `openpyxl` Table objects — filterable in Excel.

---

## Agent 6 — Opportunity Agent (`src/pso/opportunity.py`) [FUTURE]

**Responsibility:** Join city-level analysis with external population / household data
to compute market penetration, addressable volume, and coverage gaps.

**Status:** Scaffolded but not implemented. Activates when user provides the external
dataset. Schema contract defined in `config.py → OPPORTUNITY_SCHEMA`.

**Expected external data columns:**
- `CityNorm` (must match Ingest Agent normalization)
- `Population`
- `Households`
- `HouseholdsWithVehicles` (optional)
- `RegisteredVehicles` (optional)

---

## Agent 7 — Business Category Selector (`src/pso/main.py:run-category`, `src/pso/org_report.py`)

**Responsibility:** Let the user pick one Sales Org (+ Category, for Retail Business)
and run only the reports relevant to that selection, into a dedicated output folder —
without touching `pso.main run` or `workspace/run_all_reports.py`.

**Taxonomy** (`config.SALES_ORG_CATEGORIES`, hardcoded deliberately — `ingest.py`
already fails loudly on an unmapped Org/product, so a new one would surface there
before ever reaching the selector):
- **Retail Business** → `["Fuels", "Lubricants"]` — the only Org with a real
  sub-category split, and the only one that gets full deep-dive report suites.
- Every other Org (Head Office, Strategic Accounts, Agency, Key Accounts, LPG,
  Aviation, Chemicals, Power Projects, Marine) → `None` — one simple report, no
  category prompt.

**Dispatch** (`pso.main run-category --org ... [--category ...]`):
- **Retail Business + Lubricants** → `lubes_analyze.run_lubes()` + a Lubricants-only
  Excel workbook + the 9 existing `workspace/lubes_*`/`city_profiles*`/
  `national_vol_slide.py`/`frameworks.py` scripts (same list as `run_all_reports.py`).
- **Retail Business + Fuels** → `analyze.run_all()` + `premium_fuel_analyze.run_premium_fuel()`
  + a Fuels-only Excel workbook (Diesel/Petrol/Premium Fuel sheets) + the 9
  `workspace/fuels_*.py` scripts — full mirrors of the Lubricants suite, adapted to
  Diesel/Petrol/R95/PMG (R95-vs-PMG sections reuse `premium_fuel_analyze.py` directly
  rather than recomputing).
- **Any other Org** → `org_report.py` builds one small Excel overview (+ by-product
  table if >1 product code, + by-region table if >1 region) and one short Word
  summary — no deep dive, these channels have no further category worth splitting.

**Output folder:** `reports/<Org_sanitized>[_<Category>]/`. Every `workspace/fuels_*.py`
and `workspace/lubes_*.py` script routes its output there via the `PSO_OUTDIR` env var
(read by `_pso_common.out_path()`, same mechanism `PSO_INPUT` already uses for the
source file) — no per-script edits needed to add a new category folder.

**Exception — legacy hardcoded output folders:** `city_profiles.py` and
`city_profiles_volume.py` predate `PSO_OUTDIR` and hardcode their own output
directory (`reports/city_profiles/`, `reports/city_profiles_volume/`) rather than
using `out_path()`'s default. Left untouched (never edit an existing report script —
see Rules below). `main.py`'s `_run_workspace_scripts()` instead snapshots those
folders' file mtimes before running either script and, after it exits, moves any
new/changed file into `<out_dir>/city_profiles[_volume]/` — see
`_LEGACY_OUTPUT_OVERRIDES`. Any *new* script added to the selector must use
`out_path()` normally; this override map should never need another entry.

**Rules:**
- Never edit an existing report-generating script when adding a new category — each
  category gets its own new file(s); only the CLI dispatcher and `_pso_common.py`'s
  `out_path()` are shared, and both are purely additive/backward-compatible.
- The `fuels_*.py` "category" dimension is `FuelSegment` (Diesel, Petrol, Other Fuels,
  LPG) — mirrors the 4-5 lube categories' treatment; R95 vs PMG is nested inside
  Petrol via `premium_fuel_analyze.py`, not treated as a 5th sibling category (that
  would double-count portfolio share).

---

## Running the Pipeline

```bash
# Full run — ingest → analyze → lubes → premium fuel → AI insights → Excel report
uv run python -m pso.main run --input "data/input/Working File Retail Fuels Data.xlsx"

# Skip AI (faster, no API key needed)
uv run python -m pso.main run --input "..." --no-ai

# Lubes only
uv run python -m pso.main lubes --input "..."

# Data quality report only
uv run python -m pso.main quality --input "..."

# List Sales Org / Category combinations with live row counts
uv run python -m pso.main categories --input "..."

# Run the full report suite for one Org (+ Category, for Retail Business) —
# see "Agent 7 — Business Category Selector" below
uv run python -m pso.main run-category --org "Retail Business" --category "Fuels" --input "..."
uv run python -m pso.main run-category --org "Aviation" --input "..."
```

---

## Key Invariants (apply to all agents)

1. No agent modifies the source file.
2. No agent invents numbers. If a metric cannot be computed from the data it is `NaN`.
3. All intermediate DataFrames are kept in memory and not written to disk except the
   final Excel report.
4. Every Excel cell that shows a percentage change is computed from `CY` and `SPLY`
   columns in that same table — never full-year `LY`. CY only ever covers a partial
   year (e.g. 10 months of FY26), so comparing it to 12-month LY is an unequal-period
   comparison; SPLY (Same Period Last Year — the same 10-month window in the prior FY,
   derived by `pso.ingest` from the source's `%SPLY*` columns) is the only valid
   like-for-like basis. The canonical CY→SPLY column mapping lives in `pso/period.py`.
   Full-year LY columns exist in the source/derived data but are not surfaced in any
   report.
5. Period label (`10M_FY26`) is read from the sheet name and shown on every Excel sheet
   header — so reports are self-describing when the file is opened months later.
