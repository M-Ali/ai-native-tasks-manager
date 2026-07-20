"""Analysis Agent — all aggregations. No AI, no narrative, pure data."""

from __future__ import annotations

import pandas as pd
import numpy as np

from pso.config import (
    COL_ORG, COL_PRODUCT, COL_REGION, COL_CITY_NORM, COL_CORP_GRP, COL_CUST_NAME,
    COL_FUEL_SEG, COL_LUBE_CAT, COL_IS_RETAIL, COL_IS_INTL,
    COL_GRS_CY, COL_VOL_CY, COL_DISC_CY, COL_PMGN_CY, COL_NMGN_CY,
    COL_GRS_SPLY, COL_VOL_SPLY, COL_DISC_SPLY, COL_PMGN_SPLY, COL_NMGN_SPLY,
    RETAIL_ORG,
)

# ── helpers ───────────────────────────────────────────────────────────────────

def _pct_chg(cy: float | pd.Series, ly: float | pd.Series) -> float | pd.Series:
    """Safe YoY % change. Returns NaN where LY == 0."""
    if isinstance(cy, pd.Series):
        return (cy - ly).div(ly.abs().replace(0, np.nan)) * 100
    return (cy - ly) / abs(ly) * 100 if ly else np.nan


def _agg(df: pd.DataFrame, by: list[str], label_map: dict | None = None) -> pd.DataFrame:
    """Group and compute standard KPIs. Prior-period basis is SPLY (Same Period
    Last Year), not full-year LY — CY only ever covers a partial year, so it
    must be compared to an equal-length prior window."""
    g = (
        df.groupby(by, dropna=False)
        .agg(
            Rows       = (COL_VOL_CY,   "count"),
            GRS_CY     = (COL_GRS_CY,   "sum"),
            GRS_SPLY   = (COL_GRS_SPLY, "sum"),
            Vol_CY_L   = (COL_VOL_CY,   "sum"),
            Vol_SPLY_L = (COL_VOL_SPLY, "sum"),
            Disc_CY    = (COL_DISC_CY,  "sum"),
            Disc_SPLY  = (COL_DISC_SPLY,"sum"),
            PMgn_CY    = (COL_PMGN_CY,  "sum"),
            PMgn_SPLY  = (COL_PMGN_SPLY,"sum"),
            NMgn_CY    = (COL_NMGN_CY,  "sum"),
            NMgn_SPLY  = (COL_NMGN_SPLY,"sum"),
        )
        .reset_index()
    )
    g["GRS_CY_B"]      = g["GRS_CY"]     / 1e9
    g["GRS_SPLY_B"]    = g["GRS_SPLY"]   / 1e9
    g["Vol_CY_ML"]     = g["Vol_CY_L"]   / 1e6
    g["Vol_SPLY_ML"]   = g["Vol_SPLY_L"] / 1e6
    g["NMgn_CY_B"]     = g["NMgn_CY"]    / 1e9
    g["NMgn_SPLY_B"]   = g["NMgn_SPLY"]  / 1e9
    g["PMgn_CY_B"]     = g["PMgn_CY"]    / 1e9
    g["Disc_CY_B"]     = g["Disc_CY"]    / 1e9

    g["GRS_Chg_Pct"]  = _pct_chg(g["GRS_CY"],    g["GRS_SPLY"])
    g["Vol_Chg_Pct"]  = _pct_chg(g["Vol_CY_L"],  g["Vol_SPLY_L"])
    g["NMgn_Chg_Pct"] = _pct_chg(g["NMgn_CY"],   g["NMgn_SPLY"])

    # Per-litre metrics (weighted average across rows)
    g["NMgn_per_Ltr_CY"]   = np.where(g["Vol_CY_L"] > 0,   g["NMgn_CY"] / g["Vol_CY_L"], np.nan)
    g["NMgn_per_Ltr_SPLY"] = np.where(g["Vol_SPLY_L"] > 0, g["NMgn_SPLY"] / g["Vol_SPLY_L"], np.nan)
    g["Disc_per_Ltr_CY"]  = np.where(g["Vol_CY_L"] > 0, g["Disc_CY"] / g["Vol_CY_L"], np.nan)
    g["PMgn_per_Ltr_CY"]  = np.where(g["Vol_CY_L"] > 0, g["PMgn_CY"] / g["Vol_CY_L"], np.nan)

    if label_map:
        g = g.rename(columns=label_map)

    # Drop raw PKR totals — keep only scaled columns
    g = g.drop(columns=["GRS_CY", "GRS_SPLY", "Vol_CY_L", "Vol_SPLY_L",
                         "Disc_CY", "Disc_SPLY", "PMgn_CY", "PMgn_SPLY",
                         "NMgn_CY", "NMgn_SPLY"], errors="ignore")
    return g


def _vol_share(df: pd.DataFrame, vol_col: str = "Vol_CY_ML") -> pd.DataFrame:
    df = df.copy()
    total = df[vol_col].sum()
    df["Vol_Share_Pct"] = (df[vol_col] / total * 100).round(2)
    df = df.sort_values(vol_col, ascending=False).reset_index(drop=True)
    df["Vol_Cumulative_Pct"] = df["Vol_Share_Pct"].cumsum().round(2)
    return df


def _grs_share(df: pd.DataFrame, grs_col: str = "GRS_CY_B") -> pd.DataFrame:
    df = df.copy()
    total = df[grs_col].sum()
    df["GRS_Share_Pct"] = (df[grs_col] / total * 100).round(2)
    df = df.sort_values(grs_col, ascending=False).reset_index(drop=True)
    df["GRS_Cumulative_Pct"] = df["GRS_Share_Pct"].cumsum().round(2)
    return df


def _pareto(df: pd.DataFrame, vol_col: str = "Vol_CY_ML",
            grs_col: str = "GRS_CY_B") -> pd.DataFrame:
    df = _vol_share(df, vol_col)
    total_grs = df[grs_col].sum()
    df["GRS_Share_Pct"] = (df[grs_col] / total_grs * 100).round(2)
    df["GRS_Cumulative_Pct"] = df["GRS_Share_Pct"].cumsum().round(2)
    return df


# ── public API ────────────────────────────────────────────────────────────────

def run_all(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Run every analysis table. Returns dict keyed by table name."""
    tables: dict[str, pd.DataFrame] = {}

    retail = df[df[COL_IS_RETAIL]].copy()
    domestic_retail = retail[~retail[COL_IS_INTL]].copy()

    # ── Layer 1: Portfolio ─────────────────────────────────────────────────────
    tables["portfolio_summary"] = _portfolio_summary(df)

    # ── Layer 2: Retail split ─────────────────────────────────────────────────
    tables["retail_segment_split"] = _retail_segment_split(domestic_retail)

    # ── Layer 3a: Diesel ──────────────────────────────────────────────────────
    diesel = domestic_retail[domestic_retail[COL_FUEL_SEG] == "Diesel"]
    tables["diesel_by_region"]   = _agg(diesel, [COL_REGION])
    tables["diesel_by_region"]   = _vol_share(tables["diesel_by_region"])
    tables["diesel_by_city"]     = _agg(diesel, [COL_REGION, COL_CITY_NORM])
    tables["diesel_by_city"]     = _vol_share(tables["diesel_by_city"])
    tables["diesel_pareto"]      = _pareto(
        _agg(diesel, [COL_CITY_NORM]).pipe(_vol_share)
    )
    tables["diesel_by_product"]  = _agg(diesel, [COL_PRODUCT, COL_REGION])

    # ── Layer 3b: Petrol ──────────────────────────────────────────────────────
    petrol = domestic_retail[domestic_retail[COL_FUEL_SEG] == "Petrol"]
    tables["petrol_by_region"]   = _agg(petrol, [COL_REGION])
    tables["petrol_by_region"]   = _vol_share(tables["petrol_by_region"])
    tables["petrol_by_city"]     = _agg(petrol, [COL_REGION, COL_CITY_NORM])
    tables["petrol_by_city"]     = _vol_share(tables["petrol_by_city"])
    tables["petrol_pareto"]      = _pareto(
        _agg(petrol, [COL_CITY_NORM]).pipe(_vol_share)
    )
    tables["petrol_by_product"]  = _agg(petrol, [COL_PRODUCT, COL_REGION])

    # ── Layer 3c: Other Retail Fuels (SKO, FO, LPG) ──────────────────────────
    other_fuels = domestic_retail[
        domestic_retail[COL_FUEL_SEG].isin(["Other Fuels", "LPG"])
    ]
    if len(other_fuels):
        tables["other_fuels_by_region"] = _agg(other_fuels, [COL_FUEL_SEG, COL_REGION])

    # ── Layer 4: Lubricants ───────────────────────────────────────────────────
    lubes = domestic_retail[domestic_retail[COL_FUEL_SEG] == "Lubricants"]
    tables["lubes_overview"]         = _agg(lubes, [COL_LUBE_CAT]).pipe(_vol_share)
    tables["lubes_by_region"]        = _agg(lubes, [COL_REGION]).pipe(_vol_share)
    tables["lubes_category_region"]  = _agg(lubes, [COL_LUBE_CAT, COL_REGION])
    tables["lubes_by_city"]          = _agg(lubes, [COL_REGION, COL_CITY_NORM]).pipe(_vol_share)
    tables["lubes_category_city"]    = _agg(lubes, [COL_LUBE_CAT, COL_CITY_NORM])
    tables["lubes_pareto"]           = _pareto(_agg(lubes, [COL_CITY_NORM]).pipe(_vol_share))
    tables["lubes_discount_erosion"] = _lube_margin_decomp(lubes)
    tables["lubes_top_customers"]    = _top_customers(lubes, n=50)

    # ── Layer 5: Cross-dimensional ────────────────────────────────────────────
    tables["city_pareto_all"]        = _pareto(_agg(domestic_retail, [COL_CITY_NORM]))
    tables["region_performance"]     = _region_performance(domestic_retail)
    tables["underperforming_regions"] = _underperforming(tables["region_performance"])
    tables["segment_region_matrix"]  = _segment_region_matrix(domestic_retail)

    return tables


# ── private builders ──────────────────────────────────────────────────────────

def _portfolio_summary(df: pd.DataFrame) -> pd.DataFrame:
    t = _agg(df, [COL_ORG])
    total_grs = t["GRS_CY_B"].sum()
    t["GRS_Share_Pct"] = (t["GRS_CY_B"] / total_grs * 100).round(1)
    t = t.sort_values("GRS_CY_B", ascending=False).reset_index(drop=True)
    return t


def _retail_segment_split(retail: pd.DataFrame) -> pd.DataFrame:
    t = _agg(retail, [COL_FUEL_SEG])
    total_grs = t["GRS_CY_B"].sum()
    total_vol = t["Vol_CY_ML"].sum()
    t["GRS_Share_Pct"] = (t["GRS_CY_B"] / total_grs * 100).round(1)
    t["Vol_Share_Pct"] = (t["Vol_CY_ML"] / total_vol * 100).round(1)
    t = t.sort_values("GRS_CY_B", ascending=False).reset_index(drop=True)
    return t


def _lube_margin_decomp(lubes: pd.DataFrame) -> pd.DataFrame:
    """Primary Margin → Discount → Net Margin waterfall per lube category."""
    rows = []
    for cat, grp in lubes.groupby(COL_LUBE_CAT, dropna=False):
        vol_cy = grp[COL_VOL_CY].sum()
        vol_sply = grp[COL_VOL_SPLY].sum()
        pm_cy  = grp[COL_PMGN_CY].sum()
        pm_sply  = grp[COL_PMGN_SPLY].sum()
        disc_cy = grp[COL_DISC_CY].sum()
        disc_sply = grp[COL_DISC_SPLY].sum()
        nm_cy  = grp[COL_NMGN_CY].sum()
        nm_sply  = grp[COL_NMGN_SPLY].sum()

        pm_l_cy  = pm_cy  / vol_cy if vol_cy else np.nan
        pm_l_sply  = pm_sply  / vol_sply if vol_sply else np.nan
        disc_l_cy = disc_cy / vol_cy if vol_cy else np.nan
        disc_l_sply = disc_sply / vol_sply if vol_sply else np.nan
        nm_l_cy  = nm_cy  / vol_cy if vol_cy else np.nan
        nm_l_sply  = nm_sply  / vol_sply if vol_sply else np.nan

        rows.append({
            "LubeCategory":         cat or "BLANK",
            "Vol_CY_ML":            round(vol_cy / 1e6, 3),
            "Vol_SPLY_ML":          round(vol_sply / 1e6, 3),
            "Vol_Chg_Pct":          round(_pct_chg(vol_cy, vol_sply), 1) if vol_sply else np.nan,
            "PrimaryMgn_CY_B":      round(pm_cy  / 1e9, 3),
            "PrimaryMgn_SPLY_B":    round(pm_sply  / 1e9, 3),
            "PMgn_per_Ltr_CY":      round(pm_l_cy,  3) if not np.isnan(pm_l_cy)  else np.nan,
            "PMgn_per_Ltr_SPLY":    round(pm_l_sply,  3) if not np.isnan(pm_l_sply)  else np.nan,
            "Discount_CY_B":        round(disc_cy / 1e9, 3),
            "Discount_SPLY_B":      round(disc_sply / 1e9, 3),
            "Disc_per_Ltr_CY":      round(disc_l_cy, 3) if not np.isnan(disc_l_cy) else np.nan,
            "Disc_per_Ltr_SPLY":    round(disc_l_sply, 3) if not np.isnan(disc_l_sply) else np.nan,
            "NetMargin_CY_B":       round(nm_cy  / 1e9, 3),
            "NetMargin_SPLY_B":     round(nm_sply  / 1e9, 3),
            "NMgn_per_Ltr_CY":      round(nm_l_cy, 3) if not np.isnan(nm_l_cy) else np.nan,
            "NMgn_per_Ltr_SPLY":    round(nm_l_sply, 3) if not np.isnan(nm_l_sply) else np.nan,
            "NMgn_per_Ltr_Chg_Pct": round(_pct_chg(nm_l_cy, nm_l_sply), 1) if (not np.isnan(nm_l_cy) and not np.isnan(nm_l_sply) and nm_l_sply != 0) else np.nan,
            "Disc_Compression_Flag": disc_l_cy > disc_l_sply if (not np.isnan(disc_l_cy) and not np.isnan(disc_l_sply)) else False,
            "Margin_Compression_Flag": nm_l_cy < nm_l_sply if (not np.isnan(nm_l_cy) and not np.isnan(nm_l_sply)) else False,
        })

    return pd.DataFrame(rows)


def _top_customers(lubes: pd.DataFrame, n: int = 50) -> pd.DataFrame:
    t = (
        lubes.groupby([COL_CUST_NAME, COL_REGION, COL_LUBE_CAT], dropna=False)
        .agg(
            Vol_CY_ML   = (COL_VOL_CY, lambda x: x.sum() / 1e6),
            Vol_SPLY_ML = (COL_VOL_SPLY, lambda x: x.sum() / 1e6),
            NMgn_CY_B   = (COL_NMGN_CY, lambda x: x.sum() / 1e9),
        )
        .reset_index()
    )
    t["Vol_Chg_Pct"] = _pct_chg(t["Vol_CY_ML"], t["Vol_SPLY_ML"])
    # Collapse to customer level (sum across categories), keep top N
    top = (
        t.groupby(COL_CUST_NAME, dropna=False)
        .agg(Total_Vol_CY_ML=("Vol_CY_ML", "sum"),
             Total_NMgn_CY_B=("NMgn_CY_B", "sum"))
        .sort_values("Total_Vol_CY_ML", ascending=False)
        .head(n)
        .reset_index()
    )
    # Merge back to get region
    region_map = lubes.groupby(COL_CUST_NAME)[COL_REGION].agg(
        lambda x: x.mode()[0] if len(x) else ""
    )
    top[COL_REGION] = top[COL_CUST_NAME].map(region_map)
    return top


def _region_performance(retail: pd.DataFrame) -> pd.DataFrame:
    t = _agg(retail, [COL_REGION, COL_FUEL_SEG])
    total_vol = retail[COL_VOL_CY].sum()
    t["Vol_National_Share_Pct"] = (t["Vol_CY_ML"] * 1e6 / total_vol * 100).round(2)
    return t.sort_values([COL_REGION, "Vol_CY_ML"], ascending=[True, False]).reset_index(drop=True)


def _underperforming(region_perf: pd.DataFrame) -> pd.DataFrame:
    mask = (region_perf["Vol_Chg_Pct"] < 0) & (region_perf["NMgn_Chg_Pct"] < 0)
    return region_perf[mask].copy().reset_index(drop=True)


def _segment_region_matrix(retail: pd.DataFrame) -> pd.DataFrame:
    """Pivot: rows=Region, cols=FuelSegment, values=Vol_CY_ML."""
    t = _agg(retail, [COL_REGION, COL_FUEL_SEG])
    pivot = t.pivot_table(
        index=COL_REGION,
        columns=COL_FUEL_SEG,
        values="Vol_CY_ML",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    pivot.columns.name = None
    return pivot
