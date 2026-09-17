"""Write reports to workspace/out without ever overwriting an earlier version."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from autopulse.monitor import MonitorReport

OUT_DIR = Path("workspace/out")


def next_version(stem: str, suffix: str, out_dir: Path = OUT_DIR) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    n = 1
    while (path := out_dir / f"{stem}-v{n}{suffix}").exists():
        n += 1
    return path


def write_monitor(report: MonitorReport, source_file: str, out_dir: Path = OUT_DIR) -> Path:
    path = next_version(f"market_monitor_{report.month}", ".xlsx", out_dir)
    about = pd.DataFrame({
        "item": ["month", "source file", *[f"note {i + 1}" for i in range(len(report.notes))]],
        "value": [str(report.month), source_file, *report.notes],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        report.brands.to_excel(xw, sheet_name="Brands", index=False)
        report.segments.to_excel(xw, sheet_name="Segments", index=False)
        report.models.to_excel(xw, sheet_name="Models", index=False)
        about.to_excel(xw, sheet_name="About", index=False)
    return path
