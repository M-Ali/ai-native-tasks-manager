"""Read the prequalification PDF and slice it into named sections.

The EPADS-generated PDF keeps most of its body text inside XForm objects that
pypdf cannot decode -- it returns page footers and little else. Poppler's
``pdftotext`` reads the same file cleanly, so it is used when available and
pypdf is the fallback.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

DATA_DIR = Path("data")

# Headings as they appear in the document body (the same strings also appear in
# the table of contents, which is why we start searching after the invitation).
BODY_START = "INVITATION FOR PRE-QUALIFICATION"
SECTION_MARKERS: tuple[tuple[str, str], ...] = (
    ("invitation", BODY_START),
    ("instructions", "Instructions to Applicants"),
    # ITA 6.3 also names the PDS, so anchor on the sheet's own opening sentence.
    ("pds", "The following specific data for the Prequalification of Applicants"),
    ("eligibility", "Eligibility & Qualification Criteria"),
    ("evaluation", "Evaluation Criteria"),
    ("annexures", "Annexure I"),
    ("forms", "Additional Forms and Documents"),
)

_FOOTER = re.compile(
    r"e-Pak Acquisition and Disposal System \(EPADS\) generated document.*?"
    r"Page \d+ (?:of|/) \d+",
    re.DOTALL,
)


class ExtractionError(RuntimeError):
    pass


@dataclass
class Document:
    path: Path
    pages: list[str]
    engine: str

    @property
    def text(self) -> str:
        return "\n".join(self.pages)

    def clean_text(self) -> str:
        return _FOOTER.sub("", self.text)

    def sections(self) -> dict[str, str]:
        """Split the cleaned text on the known headings, in document order."""
        body = self.clean_text()
        offset = body.find(BODY_START)
        offset = 0 if offset == -1 else offset

        hits: list[tuple[int, str]] = []
        for key, marker in SECTION_MARKERS:
            idx = body.find(marker, offset)
            if idx != -1:
                hits.append((idx, key))
        hits.sort()

        out: dict[str, str] = {}
        for n, (start, key) in enumerate(hits):
            end = hits[n + 1][0] if n + 1 < len(hits) else len(body)
            out[key] = body[start:end].strip()
        return out


def find_pdf(explicit: Path | None = None, data_dir: Path = DATA_DIR) -> Path:
    """Return the PQ PDF: the explicit path, else the only PDF under data/."""
    if explicit is not None:
        if not explicit.exists():
            raise FileNotFoundError(f"No such PDF: {explicit}")
        return explicit
    candidates = sorted(p for p in data_dir.glob("*") if p.suffix.lower() == ".pdf")
    if not candidates:
        raise FileNotFoundError(f"No PDF found in {data_dir.resolve()}")
    return candidates[0]


def _pdftotext(path: Path) -> list[str] | None:
    exe = shutil.which("pdftotext")
    if exe is None:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.txt"
        result = subprocess.run(
            [exe, "-layout", str(path), str(out)],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0 or not out.exists():
            return None
        # pdftotext separates pages with a form feed.
        return out.read_text(encoding="utf-8", errors="replace").split("\f")


def _pypdf(path: Path) -> list[str]:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    return [(page.extract_text() or "") for page in reader.pages]


def load(path: Path, engine: str = "auto") -> Document:
    """Extract page text. engine: auto | pdftotext | pypdf."""
    if engine not in ("auto", "pdftotext", "pypdf"):
        raise ValueError(f"Unknown engine {engine!r}")

    if engine in ("auto", "pdftotext"):
        pages = _pdftotext(path)
        if pages is not None:
            return Document(path=path, pages=pages, engine="pdftotext")
        if engine == "pdftotext":
            raise ExtractionError(
                "pdftotext is not on PATH (install poppler-utils) -- use --engine pypdf"
            )

    return Document(path=path, pages=_pypdf(path), engine="pypdf")


def key_dates(doc: Document) -> dict[str, str]:
    """Pull the dates the PDS states, so they can be checked against criteria.Tender."""
    body = doc.clean_text()
    out: dict[str, str] = {}
    m = re.search(
        r"Clarification Date:\s*([A-Za-z]+day,\s*[A-Za-z]+ \d{1,2}, \d{4})", body
    )
    if m:
        out["clarification"] = m.group(1)
    m = re.search(
        r"Deadline for Application Submission:.*?Date:\s*"
        r"([A-Za-z]+day,\s*[A-Za-z]+ \d{1,2}, \d{4})\s*Time:\s*([\d: ]+[AP]M)",
        body,
        re.DOTALL,
    )
    if m:
        out["submission"] = f"{m.group(1)} {m.group(2).strip()}"
    return out
