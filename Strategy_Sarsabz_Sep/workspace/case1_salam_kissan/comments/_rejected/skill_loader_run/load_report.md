# Comment load report

- Source: `comments.csv`
- Rows in file: **1573**
- Non-empty cells: **9438**
- Interface scaffolding discarded (Reply / Hide replies / @handle / timestamps / bare counts): **3177**
- Too short to code (< 2 chars, mostly emoji-only): **39**
- **Comments analysed: 6219** (4646 of them replies)
- Rows whose text repeats elsewhere in the file: **4984** across 101 distinct text(s)

## Script mix

| Script | Comments |
|---|---:|
| latin | 6020 |
| other | 111 |
| urdu_script | 54 |
| mixed | 34 |

Quote **comments analysed**, never the raw row or cell count - the difference between them is interface scaffolding, not consumer voice.

## Repeated text

`times_in_file` counts how often each exact text appears. Exports repeat a comment when it shows in several reply panes, so **N copies is usually one person, not N people** - check the rows before treating a repeat as corroboration. A short phrase many people genuinely type is a real signal; a long identical sentence is almost always the same comment captured twice.

A header row was detected and excluded.
