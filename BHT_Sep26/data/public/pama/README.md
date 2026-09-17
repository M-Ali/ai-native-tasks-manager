# Monthly sales files

Put the monthly sales CSV here. Columns (one row per model per month):

| column  | example        | notes                                       |
|---------|----------------|---------------------------------------------|
| month   | 2026-08        | YYYY-MM                                     |
| brand   | Suzuki         |                                             |
| model   | Swift          |                                             |
| segment | Hatchback      | keep segment names consistent across months |
| units   | 1234           | whole number, >= 0                          |
| source  | PAMA Aug 2026  | where the figure came from                  |

Run `uv run autopulse validate <file>` before anything else.
