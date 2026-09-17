# HISTORY

Newest first.

## 2026-09-17 — v0.1.0: project created

- Set up the uv app `autopulse` (Python 3.12, pandas, openpyxl, pytest).
- `validate`: checks the monthly model-level sales CSV. It covers required
  columns, YYYY-MM months, whole units >= 0, no blank fields and no duplicate
  month/brand/model rows, and it reports every problem with its spreadsheet
  row number.
- `monitor`: works out units, market and segment share, month-on-month and
  year-on-year change, and brand share change in points. It writes a versioned
  xlsx to `workspace/out/`.
- Split the data folders into `data/public/` and git-ignored
  `data/clients/<client>/`, because the group handles competing brands.
- Added AGENTS.md, README.md, HISTORY.md and SOFAR.md.
- No real market data yet. The tests use made-up numbers.
