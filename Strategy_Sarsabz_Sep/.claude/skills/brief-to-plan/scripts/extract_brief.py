"""Extract a brief and report what it references but does not contain.

Usage:
    python extract_brief.py <brief.pdf|brief.docx> --out <dir> [--engine auto|pdftotext|pypdf]

Writes:
    full.txt              cleaned full text
    sections.md           the document split on its own headings
    missing_sections.md   cross-references with no matching heading  <- read this first
    dates.md              dates found, labelled where possible

Prints characters-per-page so a silently half-extracted document is visible
immediately. See references/extraction-gotchas.md.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

# Boilerplate footers that portal-generated PDFs repeat on every page.
FOOTER_PATTERNS = [
    r"e-Pak Acquisition and Disposal System \(EPADS\) generated document.*?Page \d+ (?:of|/) \d+",
    r"Page \d+ of \d+",
]

# Headings a brief is likely to be built from. Matched at line start only.
HEADING_RE = re.compile(
    r"^\s*("
    r"(?:SECTION|Section)\s+(?:[IVXLC]+|\d+)"
    r"|(?:ANNEXURE|Annexure|ANNEX|Annex|APPENDIX|Appendix)\s*[-–]?\s*(?:[IVXLC]+|\d+|[A-Z])"
    r"|(?:SCHEDULE|Schedule)\s+of\s+\w+"
    r"|(?:PART|Part)\s+(?:[IVXLC]+|\d+)"
    r")\b",
    re.MULTILINE,
)

# The same names appearing anywhere, including mid-sentence cross-references.
REFERENCE_RE = re.compile(
    r"(?:Section|SECTION)\s+(?:[IVXLC]+|\d+)"
    r"|(?:Annexure|ANNEXURE|Annex|ANNEX|Appendix|APPENDIX)\s*[-–]?\s*(?:[IVXLC]+|\d+|[A-Z])\b"
    r"|(?:Schedule|SCHEDULE)\s+of\s+[A-Z]\w+"
)

DATE_RE = re.compile(
    r"(?:[A-Z][a-z]+day,\s*)?[A-Z][a-z]+ \d{1,2},? \d{4}(?:\s+\d{1,2}:\d{2}\s*[AP]M)?"
    r"|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r"|\d{4}-\d{2}-\d{2}"
)

DATE_LABELS = [
    "deadline", "submission", "closing", "clarification", "opening", "pre-bid",
    "pre-application", "start date", "end date", "completion", "validity", "award",
]


def _pdftotext(path: Path) -> list[str] | None:
    exe = shutil.which("pdftotext")
    if exe is None:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.txt"
        res = subprocess.run([exe, "-layout", str(path), str(out)], capture_output=True, text=True)
        if res.returncode != 0 or not out.exists():
            return None
        return out.read_text(encoding="utf-8", errors="replace").split("\f")


def _pypdf(path: Path) -> list[str]:
    from pypdf import PdfReader

    return [(p.extract_text() or "") for p in PdfReader(str(path)).pages]


def _docx(path: Path) -> list[str]:
    """Walk the body so tables come through in order -- briefs put criteria in tables."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))
    out: list[str] = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = Paragraph(child, doc).text.strip()
            if text:
                out.append(text)
        elif child.tag == qn("w:tbl"):
            for row in Table(child, doc).rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                out.append(" | ".join(cells))
    return ["\n".join(out)]


def extract(path: Path, engine: str) -> tuple[list[str], str]:
    if path.suffix.lower() == ".docx":
        return _docx(path), "python-docx"
    if engine in ("auto", "pdftotext"):
        pages = _pdftotext(path)
        if pages is not None:
            return pages, "pdftotext"
        if engine == "pdftotext":
            raise SystemExit("pdftotext not on PATH -- install poppler or use --engine pypdf")
    return _pypdf(path), "pypdf"


def clean(text: str) -> str:
    for pattern in FOOTER_PATTERNS:
        text = re.sub(pattern, "", text, flags=re.DOTALL)
    return text


def find_headings(text: str) -> list[tuple[int, str]]:
    hits = []
    for m in HEADING_RE.finditer(text):
        line = text[m.start():text.find("\n", m.start()) if "\n" in text[m.start():] else len(text)]
        # A heading is a short line. A long line is prose that happens to start with one.
        if len(line.strip()) <= 90:
            hits.append((m.start(), " ".join(m.group(1).split())))
    return hits


def normalise(ref: str) -> str:
    """Canonical form so 'Annexure - II', 'ANNEXURE II' and 'Annexure II' all match."""
    ref = " ".join(ref.split()).replace("–", "-").replace("—", "-")
    ref = re.sub(r"\s*-\s*", " ", ref).rstrip(".,;:")
    parts = ref.split()
    if not parts:
        return ref
    head = parts[0].capitalize()
    tail = [p.upper() if re.fullmatch(r"[ivxlcIVXLC]+|[A-Za-z]|\d+", p) else p.lower()
            for p in parts[1:]]
    return " ".join([head, *tail])


def label_for(text: str, ref: str) -> str:
    """The descriptive name a cross-reference gives, e.g. Section VII -> 'Schedule of...'."""
    for m in re.finditer(re.escape(ref) + r"\s*[-–—:(]\s*([A-Z][A-Za-z &/,']{3,60})",
                         text, re.IGNORECASE):
        return " ".join(m.group(1).split()).rstrip(")., ")
    return ""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("brief", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--engine", default="auto", choices=["auto", "pdftotext", "pypdf"])
    args = ap.parse_args()

    if not args.brief.exists():
        raise SystemExit(f"No such file: {args.brief}")
    args.out.mkdir(parents=True, exist_ok=True)

    pages, engine = extract(args.brief, args.engine)
    print(f"{args.brief.name}: {len(pages)} page(s) via {engine}")

    sizes = [len(p) for p in pages]
    if len(pages) > 1:
        print("chars/page: " + " ".join(str(s) for s in sizes))
        small = [i + 1 for i, s in enumerate(sizes) if s < 200]
        if len(small) > len(pages) / 3:
            print(
                f"  WARNING: {len(small)} page(s) under 200 chars -- likely headers/footers only.\n"
                f"  See references/extraction-gotchas.md before trusting this output."
            )

    text = clean("\n".join(pages))
    (args.out / "full.txt").write_text(text, encoding="utf-8")
    print(f"  {len(text)} characters extracted")

    # --- sections ---------------------------------------------------------
    headings = find_headings(text)
    lines = ["# Sections found", ""]
    if headings:
        lines += ["| Heading | Characters |", "|---|---:|"]
        for n, (start, name) in enumerate(headings):
            end = headings[n + 1][0] if n + 1 < len(headings) else len(text)
            lines.append(f"| {name} | {end - start} |")
        lines += [
            "",
            "A section of a few dozen characters is a false hit (usually the table of",
            "contents). One of tens of thousands has probably swallowed its neighbour.",
            "",
        ]
        for n, (start, name) in enumerate(headings):
            end = headings[n + 1][0] if n + 1 < len(headings) else len(text)
            lines += [f"## {name}", "", "```", text[start:end].strip()[:4000], "```", ""]
    else:
        lines.append("No headings matched. The document may use its own naming - split by hand.")
    (args.out / "sections.md").write_text("\n".join(lines), encoding="utf-8")

    # --- referenced but absent -------------------------------------------
    heading_names = {normalise(name) for _, name in headings}
    referenced = {normalise(m.group(0)) for m in REFERENCE_RE.finditer(text)}
    missing = sorted(referenced - heading_names)

    lines = ["# Referenced but no matching heading found", ""]
    if missing:
        lines += [
            f"**{len(missing)} cross-reference(s)** point at something this document does not",
            "appear to contain. Some are harmless naming mismatches. Some are the whole ask.",
            "Verify each by hand, then check the portal listing for attachments the main file",
            "does not include.",
            "",
            "| Referenced | Named as | Times | Context |",
            "|---|---|---:|---|",
        ]
        counted = []
        for ref in missing:
            hits = list(re.finditer(re.escape(ref), text, re.IGNORECASE))
            counted.append((len(hits), ref, hits))
        for n, ref, hits in sorted(counted):
            m = hits[0] if hits else None
            context = " ".join(text[max(0, m.start() - 80):m.end() + 80].split()) if m else ""
            lines.append(f"| {ref} | {label_for(text, ref) or '-'} | {n} | ...{context}... |")
        lines += [
            "",
            "Fewest mentions first - a reference made once, in passing, is the likeliest to be",
            "something the buyer forgot to attach.",
            "",
            "Numbered sections are often carried in the body under descriptive headings instead.",
            "Reconcile against the headings that were found:",
            "",
        ]
        lines += [f"- {n}" for n in dict.fromkeys(name for _, name in headings)] or ["- (none)"]
    else:
        lines.append("Every cross-reference resolves to a heading in the document.")
    lines += ["", "Also check: addenda issued since this version, and attachments on the portal."]
    (args.out / "missing_sections.md").write_text("\n".join(lines), encoding="utf-8")

    # --- dates ------------------------------------------------------------
    lines = ["# Dates found", "", "| Date | Context |", "|---|---|"]
    seen: set[str] = set()
    labelled = 0
    for m in DATE_RE.finditer(text):
        value = m.group(0).strip()
        context = " ".join(text[max(0, m.start() - 110):m.start()].split())[-110:]
        key = f"{value}|{context[-40:]}"
        if key in seen:
            continue
        seen.add(key)
        hit = any(lbl in context.lower() for lbl in DATE_LABELS)
        labelled += hit
        mark = "**" if hit else ""
        lines.append(f"| {mark}{value}{mark} | ...{context} |")
    lines += ["", "Bold rows sit near a deadline-ish word. Confirm every one against the",
              "document itself before putting it in a plan."]
    (args.out / "dates.md").write_text("\n".join(lines), encoding="utf-8")

    print(f"  {len(headings)} heading(s), {len(missing)} unresolved reference(s), "
          f"{len(seen)} date(s) ({labelled} near a deadline word)")
    print(f"Wrote {args.out}/full.txt, sections.md, missing_sections.md, dates.md")
    if missing:
        print("  -> read missing_sections.md first")


if __name__ == "__main__":
    main()
