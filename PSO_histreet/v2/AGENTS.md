# Pricing Intelligence v2 — AGENTS.md

## Product Vision
A config-driven, platform-agnostic pricing strategy engine.
Phase 1: PSO Lubricants on Daraz.pk.
Future: Any product category on any e-commerce platform.

---

## Architecture

```
v2/
  config/
    lubricants.yaml     ← ALL brands, grades, platforms defined here (no code changes)
  scrapers/
    base.py             ← Abstract BaseScraper + ScrapedProduct dataclass
    daraz.py            ← Daraz.pk Playwright implementation
  db/
    store.py            ← SQLite persistence with timestamped runs
    prices.db           ← Auto-created on first scrape
  normalizer.py         ← Extracts grade, pack size, brand from raw titles
  analyzer.py           ← 5 charts + console gap table
  reporter.py           ← Word document with recommendations
  main.py               ← CLI: scrape / analyze / report / all
  output/
    charts/             ← PNG charts (01-05)
    reports/            ← PSO_Pricing_Intelligence_Report.docx
```

---

## How to Run

```bash
# Full pipeline (scrape → analyze → report)
uv run python main.py all

# Individual steps
uv run python main.py scrape          # scrape Daraz only
uv run python main.py scrape --brands "Shell,ZIC"   # specific brands
uv run python main.py scrape --dry-run               # preview search terms
uv run python main.py analyze         # charts from DB
uv run python main.py report          # Word doc from DB
```

---

## Analysis Modules

| Chart | File | What It Shows |
|-------|------|---------------|
| 01 | `01_price_matrix_pcmo/hdeo.png` | Box plot: Price/L per brand per grade |
| 02 | `02_tier_gaps.png` | PSO vs market median by tier (Super Premium → Economy) |
| 03 | `03_pack_premium_curve.png` | Price/L vs pack size — should decline as size grows |
| 04 | `04_price_index.png` | PSO = 100 index; all competitors relative to PSO per grade |
| 05 | `05_price_heatmap.png` | Full Brand × Grade heatmap of median price/L |

---

## Config-Driven Design

To analyse a **new product category** (e.g. tyres, batteries):
1. Create `config/tyres.yaml` with the same schema
2. Pass `--config config/tyres.yaml` (future flag) or copy over `lubricants.yaml`
3. Zero code changes needed

To add a **new platform** (e.g. OLX, Amazon.ae):
1. Create `scrapers/olx.py` extending `BaseScraper`
2. Register in `scrapers/__init__.py` PLATFORM_MAP
3. Add to `config/lubricants.yaml` platforms section

---

## Data Model (SQLite)

| Column | Description |
|--------|-------------|
| platform | Source (daraz) |
| brand_query | Search term used |
| title | Raw product title |
| price | PKR |
| original_price | Pre-discount PKR (if available) |
| brand_detected | Normalised brand name |
| grade_detected | e.g. 5W-30 |
| pack_size_l | Litres (extracted from title) |
| price_per_litre | price / pack_size_l |
| oil_type | pcmo / hdeo / mco |
| scraped_at | UTC timestamp |

Price history is cumulative — every scrape run adds new rows.
Use `db/store.py::all_runs_summary()` to query price trends over time.

---

## Extending the Pipeline (Roadmap)

- [ ] Add `--config` CLI flag to switch product categories
- [ ] OLX.com.pk scraper (static HTML, simpler than Daraz)
- [ ] Brand website scrapers (Shell PK, TotalEnergies PK)
- [ ] Price trend charts (time series from multi-run DB)
- [x] Excel export of price matrix alongside Word report — **DONE** (see below)
- [ ] Claude API integration for AI-generated pricing narrative
- [ ] Email/Slack delivery of weekly price alert digest

---

## Deliverables Log

### 2026-07-08 — PSO Pricing Strategy PPTX v4
- **File:** `output/reports/PSO_Pricing_Strategy_v4.pptx`
- 20 slides covering all 8 pricing frameworks (F1-F8) with speaker notes
- Fixed F5 spec premium chart to clustered bar (PSO vs Market Median, two-series)
- Built with pptxgenjs via `pptx_workspace/create_pptx.js`

### 2026-07-08/09 — PSO Pricing Strategy Excel Workbook v2
- **File:** `output/reports/PSO_Pricing_Strategy_Workbook_v2.xlsx`
- **Builder:** `build_workbook_v2.py` (clean rebuild — pure ASCII strings, zero encoding issues)
- 11 sheets: README, Competitor Data, F1-F8, Summary
- 45 PSO SKUs x 8 frameworks; 606 live Excel formulas
- Each F-sheet: parameter tables + live calculation table + HOW IT WORKS / RATIONALE / IMPLICATIONS methodology blocks
- Competitor Data: all 365 Daraz.pk listings + grade-level market statistics
- Summary: weighted synthesis (F1 30% + F2 30% + F3 20% + F5 10% + F4 10%), F4 hard floor override, confidence rating, vs-market signal per SKU
- **Encoding fix history:** v1 builder (`build_excel.py`) was corrupted by PowerShell cp1252 read; fixed via `fix_encoding.py` (cp1252->utf-8 repair, 123 cells); v2 avoids the issue entirely by using only ASCII in source strings
