"""Self-assessment: track what the agency actually holds and score it against the 100 marks."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

from . import criteria

STATUS_PATH = Path("workspace/status.yaml")

# What a status value means when converting to marks.
CREDIT = {
    "ready": 1.0,      # evidence in hand, attested, ready to upload
    "partial": 0.5,    # exists but incomplete / unattested / needs refresh
    "missing": 0.0,    # not held
    "na": 0.0,         # not applicable to this agency (still scores nothing)
}
VALID = tuple(CREDIT)


@dataclass
class Line:
    key: str
    label: str
    bucket: str
    marks: int
    status: str
    note: str

    @property
    def earned(self) -> float:
        return round(self.marks * CREDIT.get(self.status, 0.0), 2)


@dataclass
class Assessment:
    agency: str
    lines: list[Line]
    eligibility: list[Line]

    @property
    def earned(self) -> float:
        return round(sum(line.earned for line in self.lines), 2)

    @property
    def available(self) -> int:
        return sum(line.marks for line in self.lines)

    @property
    def passes(self) -> bool:
        return self.earned >= criteria.Tender().passing_marks

    def by_bucket(self) -> dict[str, tuple[float, int]]:
        out: dict[str, tuple[float, int]] = {}
        for line in self.lines:
            got, avail = out.get(line.bucket, (0.0, 0))
            out[line.bucket] = (round(got + line.earned, 2), avail + line.marks)
        return out

    def gaps(self) -> list[Line]:
        """Scored lines losing marks, worst first."""
        losing = [ln for ln in self.lines if ln.earned < ln.marks]
        return sorted(losing, key=lambda ln: ln.marks - ln.earned, reverse=True)

    def eligibility_gaps(self) -> list[Line]:
        return [ln for ln in self.eligibility if ln.status in ("missing", "partial")]


def template(agency: str = "<your agency name>") -> dict:
    """A fresh status file with every criterion set to 'missing'."""
    return {
        "agency": agency,
        "tender": criteria.Tender().reference,
        "eligibility": {
            item.key: {"status": "missing", "note": ""} for item in criteria.ELIGIBILITY
        },
        "scored": {
            item.key: {"status": "missing", "note": ""} for item in criteria.SCORED
        },
    }


def init(path: Path = STATUS_PATH, agency: str = "<your agency name>", force: bool = False) -> Path:
    if path.exists() and not force:
        raise FileExistsError(f"{path} already exists; pass --force to overwrite")
    path.parent.mkdir(parents=True, exist_ok=True)
    header = (
        "# Self-assessment for the SLIC advertising-agency prequalification (P78118).\n"
        "# status: ready | partial | missing | na\n"
        "# 'ready' scores full marks, 'partial' scores half, the rest score nothing.\n"
    )
    path.write_text(
        header + yaml.safe_dump(template(agency), sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )
    return path


def _lines(raw: dict, items, label_of, bucket_of) -> list[Line]:
    out: list[Line] = []
    for item in items:
        entry = raw.get(item.key) or {}
        status = str(entry.get("status", "missing")).lower()
        if status not in VALID:
            raise ValueError(
                f"{item.key}: status {status!r} is not one of {', '.join(VALID)}"
            )
        out.append(
            Line(
                key=item.key,
                label=label_of(item),
                bucket=bucket_of(item),
                marks=getattr(item, "marks", 0),
                status=status,
                note=str(entry.get("note", "") or ""),
            )
        )
    return out


def load(path: Path = STATUS_PATH) -> Assessment:
    if not path.exists():
        raise FileNotFoundError(f"{path} not found -- run 'slic-pq init' first")
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    scored = _lines(
        raw.get("scored") or {},
        criteria.SCORED,
        lambda i: i.question,
        lambda i: i.bucket,
    )
    elig = _lines(
        raw.get("eligibility") or {},
        criteria.ELIGIBILITY,
        lambda i: i.requirement,
        lambda _: "eligibility",
    )
    return Assessment(agency=str(raw.get("agency", "")), lines=scored, eligibility=elig)
