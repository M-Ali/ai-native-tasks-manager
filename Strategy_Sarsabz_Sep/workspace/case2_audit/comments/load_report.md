# Comment load report

- Source: `COMMENTS.xlsx`
- Rows in file: **1531**
- Non-empty cells: **1532**
- Interface scaffolding discarded (Reply / Hide replies / @handle / timestamps / bare counts): **82**
- Too short to code (< 2 chars, mostly emoji-only): **15**
- **Comments analysed: 1433** (499 of them replies)
- Rows whose text repeats elsewhere in the file: **190** across 42 distinct text(s)

## Script mix

| Script | Comments |
|---|---:|
| latin | 1025 |
| urdu_script | 366 |
| mixed | 32 |
| other | 10 |

Quote **comments analysed**, never the raw row or cell count - the difference between them is interface scaffolding, not consumer voice.

## Repeated text

`times_in_file` counts how often each exact text appears. Exports repeat a comment when it shows in several reply panes, so **N copies is usually one person, not N people** - check the rows before treating a repeat as corroboration. A short phrase many people genuinely type is a real signal; a long identical sentence is almost always the same comment captured twice.

A header row was detected and excluded.
