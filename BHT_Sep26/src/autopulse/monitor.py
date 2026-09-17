"""Market monitor: share and growth by model, brand and segment for one month."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class MonitorReport:
    month: pd.Period
    models: pd.DataFrame
    brands: pd.DataFrame
    segments: pd.DataFrame
    notes: list[str]


def build_report(sales: pd.DataFrame, month: str | None = None) -> MonitorReport:
    """Build the monitor for `month` (YYYY-MM); defaults to the latest month in the file."""
    months = set(sales["month"])
    target = pd.Period(month, "M") if month else max(months)
    if target not in months:
        raise ValueError(f"{target} is not in the sales file")
    prev, last_year = target - 1, target - 12

    notes = []
    if prev not in months:
        notes.append(f"{prev} not in file: month-on-month change left blank")
    if last_year not in months:
        notes.append(f"{last_year} not in file: year-on-year change left blank")

    models = _compare(sales, ["segment", "brand", "model"], target, prev, last_year)
    total = models["units"].sum()
    seg_totals = models.groupby("segment")["units"].transform("sum")
    models["share_of_segment_pct"] = _pct(models["units"], seg_totals)
    models["share_of_market_pct"] = _pct(models["units"], total)
    models = models.sort_values(["segment", "units"], ascending=[True, False])

    brands = _compare(sales, ["brand"], target, prev, last_year)
    brands["share_of_market_pct"] = _pct(brands["units"], total)
    brands = _add_share_change(brands, sales, prev, last_year)
    brands = brands.sort_values("units", ascending=False)

    segments = _compare(sales, ["segment"], target, prev, last_year)
    segments["share_of_market_pct"] = _pct(segments["units"], total)
    segments = segments.sort_values("units", ascending=False)

    return MonitorReport(
        target,
        models.reset_index(drop=True),
        brands.reset_index(drop=True),
        segments.reset_index(drop=True),
        notes,
    )


def _compare(sales, keys, target, prev, last_year) -> pd.DataFrame:
    months = set(sales["month"])

    def units_in(period, name):
        return sales[sales["month"] == period].groupby(keys)["units"].sum().rename(name)

    cur = units_in(target, "units")
    out = (
        pd.concat([cur, units_in(prev, "units_prev_month"), units_in(last_year, "units_last_year")], axis=1)
        .loc[cur.index]
        .reset_index()
    )
    out["mom_change_pct"] = _growth(out["units"], out["units_prev_month"], prev in months)
    out["yoy_change_pct"] = _growth(out["units"], out["units_last_year"], last_year in months)
    return out


def _add_share_change(brands, sales, prev, last_year) -> pd.DataFrame:
    for period, col in [(prev, "share_pts_vs_prev_month"), (last_year, "share_pts_vs_last_year")]:
        base = sales[sales["month"] == period]
        if base.empty:
            brands[col] = pd.NA
            continue
        share = _pct(base.groupby("brand")["units"].sum(), base["units"].sum())
        brands[col] = (brands["share_of_market_pct"] - brands["brand"].map(share).fillna(0)).round(1)
    return brands


def _growth(cur, base, base_month_present: bool) -> pd.Series:
    # A model missing from a month that IS in the file sold zero; a month missing from the file is unknown.
    if not base_month_present:
        return pd.Series(pd.NA, index=cur.index, dtype="Float64")
    base = base.fillna(0)
    growth = (cur - base) / base.where(base > 0) * 100
    return growth.round(1).astype("Float64")


def _pct(part, whole) -> pd.Series:
    return (part / whole * 100).round(1)
