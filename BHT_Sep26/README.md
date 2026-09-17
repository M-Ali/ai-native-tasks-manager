# autopulse

Analytics tools for the group's automobile clients (Suzuki, Hyundai and others).
Built as one uv Python app. Tools that compare brands use **public data only**.
Each client's own data stays in its own folder and is never pooled.

## Status

| # | Tool | Status |
|---|------|--------|
| 1 | Market monitor: share and growth by brand, segment and model | **built** (`monitor`) |
| 2 | Used-car price and "on money" tracker (PakWheels / OLX) | planned; check terms of service first |
| 3 | Consumer voice tracker (reviews, YouTube, forums) | planned |
| 4 | Share of voice vs share of market | planned |
| 5 | Launch tracker (first 90 days) | planned |
| 6 | Demand drivers (interest rates, auto financing, exchange rate, fuel, prices) | planned |
| 7 | Leads to sales, one client at a time (GA4, CRM, dealers) | planned; private data |

## Setup

```bash
uv sync
uv run pytest
```

## Usage

1. Put the monthly model-level sales CSV in `data/public/pama/`.
   `data/public/pama/README.md` lists the columns.
2. Check it:
   ```bash
   uv run autopulse validate data/public/pama/sales.csv
   ```
3. Build the monitor (defaults to the latest month in the file):
   ```bash
   uv run autopulse monitor data/public/pama/sales.csv --month 2026-08
   ```
   This writes `workspace/out/market_monitor_2026-08-v1.xlsx` with Brands,
   Segments, Models and About sheets. A second run writes `-v2`. Earlier
   versions are never overwritten.

## What the monitor computes

- Units, share of the total market and share of each segment.
- Change vs last month and vs the same month last year, in %.
- For brands, the change in share in percentage points.
- If the file has no data for the comparison month, the change is left blank
  and the About sheet says why. If that month is in the file but a model has no
  row, the model counts as zero sales, and growth from zero is left blank.

## Layout

```
src/autopulse/     sales.py (load + validate), monitor.py, output.py, cli.py
tests/             made-up numbers only, never market data
data/public/       data every client may see
data/clients/      one folder per client; git-ignored
workspace/out/     generated reports; git-ignored
```

See `AGENTS.md` for working rules, `HISTORY.md` for changes and `SOFAR.md`
for the conversation that shaped the project.
