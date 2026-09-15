= Extraction gotchas

A half-extracted brief is more dangerous than an unextracted one — you read what came through
and never learn what did not. Check the extraction before trusting it.

## The quick sanity check

`extract_brief.py` prints characters per page. Scan it. **Pages returning roughly the same
small number (50-150 characters) are returning headers and footers only** — the body did not
come through.

If total extracted text looks small for the page count, something failed. A 34-page tender
should yield tens of thousands of characters, not fifteen thousand.

## Portal-generated PDFs

Documents produced by e-procurement systems (EPADS, and others built the same way) often place
body text inside XForm objects that pure-Python readers cannot decode. pypdf returns page
footers and nothing else, with no error — it silently gives you 15k characters where the file
holds 44k.

**Poppler's `pdftotext -layout` reads these correctly.** The extractor prefers it and falls back
to pypdf, reporting which engine ran. If you get the pypdf fallback on a portal document, check
the output before proceeding rather than assuming.

Install poppler if it is missing; on Windows the Git Bash environment often already has
`pdftotext` on PATH.

## Scanned or image-only PDFs

If extraction returns almost nothing on every page, the document is images. You need OCR, or the
original from the buyer. Ask — buyers usually have a text version and will send it, and asking
is faster than OCR quality problems.

## Table of contents versus body

Section headings appear at least twice: once in the contents, once as the heading itself, and
often a third time in a cross-reference ("as set out in Section III"). Naive slicing on first
occurrence carves the document at the table of contents and produces sections a few dozen
characters long.

The extractor starts after the first substantive heading to skip the contents. Where a heading
is also a common phrase in the body, anchor the split on the section's own opening sentence
instead. Check the reported section sizes — a "section" of 40 characters is a false hit, and a
section of 20,000 has probably swallowed its neighbour.

## Cross-references to nothing

The most valuable output of the extraction step. The script lists section numbers and named
annexes referenced in the text that have no matching heading in the document.

Some are harmless — a reference to a clause that exists under a slightly different name. Some
are the whole ballgame: a schedule of requirements that the buyer never attached, and that
contains the actual brief.

Verify each flagged item by hand. Then check the portal listing, which frequently carries
attachments the main PDF does not.

## Addenda

Procurement rules generally let a buyer amend by addendum up to shortly before the deadline, and
an addendum overrides the base document. If you extracted a document a week ago, re-check the
portal before finalising anything.

Where an addendum lands inside the last few days, the deadline usually extends — worth knowing
before you burn a weekend.

## DOCX

Simpler, with one trap: `python-docx` iterating paragraphs alone **silently drops every table**,
and briefs put their requirements, criteria and weightings in tables. Walk the document body
element by element so paragraphs and tables come through in order. `extract_brief.py` does this.

Watch for content in headers, footers and text boxes, which a body walk still misses. If a
document mentions a figure you cannot find in the extraction, look there.

## Encoding on Windows

Extracted text routinely contains characters outside the system code page — en dashes, curly
quotes, arrows. Read and write with `encoding="utf-8"` explicitly. Printing raw extracted text
to a Windows console will crash on the first one.
