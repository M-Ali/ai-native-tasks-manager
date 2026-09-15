"""Read the brief PDF: text layer where there is one, rendered images where there is not.

The Fatima brief mixes both. Pages 1-15 carry text; pages 16-26 (brand guidelines) and the
brand-health charts are images, so a text-only read silently misses the brand key, the
competitive frame and the awareness trend. `load()` records which pages are image-only and
`render()` writes them out as PNGs so they can actually be looked at.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

DATA_DIR = Path("data")
IMAGE_ONLY_CHARS = 40  # below this a page is treated as image-only


@dataclass
class Document:
    path: Path
    pages: list[str]

    @property
    def text(self) -> str:
        return "\n".join(f"--- page {i} ---\n{t}" for i, t in enumerate(self.pages, 1))

    def image_only_pages(self) -> list[int]:
        return [i for i, t in enumerate(self.pages, 1) if len(t.strip()) < IMAGE_ONLY_CHARS]


def find_pdf(explicit: Path | None = None, data_dir: Path = DATA_DIR) -> Path:
    """The explicit path, else the only PDF directly under data/."""
    if explicit is not None:
        if not explicit.exists():
            raise FileNotFoundError(f"No such PDF: {explicit}")
        return explicit
    candidates = sorted(p for p in data_dir.glob("*") if p.suffix.lower() == ".pdf")
    if not candidates:
        raise FileNotFoundError(f"No PDF found in {data_dir.resolve()}")
    if len(candidates) > 1:
        raise FileNotFoundError(f"Several PDFs in {data_dir}; pass one explicitly")
    return candidates[0]


def load(path: Path) -> Document:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    return Document(path=path, pages=[(p.extract_text() or "") for p in reader.pages])


def render(path: Path, out_dir: Path, pages: list[int] | None = None, dpi: int = 110) -> list[Path]:
    """Render pages (1-based; default all) to PNG with PyMuPDF. Poppler is not required."""
    import pymupdf

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(str(path))
    wanted = pages or list(range(1, doc.page_count + 1))
    written = []
    for n in wanted:
        target = out_dir / f"p{n:02d}.png"
        doc[n - 1].get_pixmap(dpi=dpi).save(str(target))
        written.append(target)
    return written
