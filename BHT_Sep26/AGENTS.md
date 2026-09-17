# AGENTS.md

Working rules for this repository. Read this before changing anything.

---

## What this project is

`autopulse` is a set of analytics tools a marketing communication group builds
for the automobile clients it handles, including **Suzuki and Hyundai**. The
market is Pakistan. The first tool is the monthly market monitor. The roadmap
is in `README.md`.

## The rules that matter

1. **Competing clients: keep their data apart.** The group handles brands that
   compete with each other.
   - `data/public/` holds public sources only (PAMA, listings, social, search
     trends). Any client may see outputs built from it.
   - `data/clients/<client>/` holds one client's own data. No command that
     produces a cross-brand output may read it. Never commit it.
2. **Never invent a number.** Every figure must come from a file in `data/`,
   with a `source` recorded. Tests use made-up numbers and say so. Never put
   them in `data/` or in a deliverable.
3. **Never overwrite a deliverable.** Reports go to `workspace/out/` through
   `output.next_version()`, which picks the next `-vN` filename. Never delete an
   earlier version, even as a test or rebuild step.
4. **Check the input first.** Run `autopulse validate` on a new file before
   building anything from it. Validation lists every problem at once and does
   not guess at fixes.
5. **Blank means unknown.** If a comparison month is missing from the file,
   growth is left blank and a note says so. Don't fill it with 0.

## Commands

```bash
uv sync
uv run pytest
uv run autopulse validate <sales.csv>
uv run autopulse monitor <sales.csv> [--month YYYY-MM]
```

## Adding a tool

- Put each tool in its own module under `src/autopulse/`, with a `cli.py` subcommand.
- Tests go in `tests/`, and every new calculation gets a hand-checked test.
- In the same change, update the status table in `README.md` and add an entry
  to `HISTORY.md`.

## Keeping the docs current

- `HISTORY.md`: what changed and why, newest first, with dates.
- `SOFAR.md`: the running record of the conversation with the user: what they
  asked, what was decided, open questions. Add to it after every working
  session. Don't rewrite earlier entries.

## Environment

Windows, PowerShell or Git Bash, uv, Python 3.12. This folder sits inside the
larger `D:\Personal` git repository. Commit only when the user asks.
