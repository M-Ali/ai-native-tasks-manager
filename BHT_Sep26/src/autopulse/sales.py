"""Load and validate the monthly model-level sales file."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

REQUIRED = ["month", "brand", "model", "segment", "units", "source"]


class SalesDataError(ValueError):
    """The sales file cannot be trusted as input."""


def load_sales(path: str | Path) -> pd.DataFrame:
    """Read a sales CSV and return it cleaned, or raise SalesDataError listing every problem."""
    df = pd.read_csv(path, dtype=str).rename(columns=lambda c: c.strip().lower())
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise SalesDataError(f"missing columns: {', '.join(missing)}")

    df = df[REQUIRED].copy()
    for col in ["month", "brand", "model", "segment", "source"]:
        df[col] = df[col].fillna("").str.strip()

    problems: list[str] = []
    for col in ["brand", "model", "segment", "source"]:
        blank = df.index[df[col] == ""]
        if len(blank):
            problems.append(f"blank {col} on rows {_rows(blank)}")

    period = pd.to_datetime(df["month"], format="%Y-%m", errors="coerce")
    bad_month = df.index[period.isna()]
    if len(bad_month):
        problems.append(f"month not YYYY-MM on rows {_rows(bad_month)}")

    units = pd.to_numeric(df["units"].fillna("").str.replace(",", ""), errors="coerce")
    bad_units = df.index[units.isna() | (units < 0) | (units % 1 != 0)]
    if len(bad_units):
        problems.append(f"units not a whole number >= 0 on rows {_rows(bad_units)}")

    dupes = df.index[df.duplicated(["month", "brand", "model"], keep=False)]
    if len(dupes):
        problems.append(f"duplicate month/brand/model on rows {_rows(dupes)}")

    if problems:
        raise SalesDataError("; ".join(problems))

    df["month"] = period.dt.to_period("M")
    df["units"] = units.astype(int)
    return df.sort_values(["month", "brand", "model"]).reset_index(drop=True)


def _rows(index: pd.Index) -> str:
    # +2 = header line plus 1-based numbering, so row numbers match the spreadsheet
    shown = [str(i + 2) for i in index[:10]]
    more = f" (+{len(index) - 10} more)" if len(index) > 10 else ""
    return ", ".join(shown) + more
