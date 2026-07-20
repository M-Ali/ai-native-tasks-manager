"""Lubes Agent — dedicated lubricants deep-dive analysis."""

from __future__ import annotations

import pandas as pd
import numpy as np

from pso.config import (
    COL_PRODUCT, COL_REGION, COL_CITY_NORM, COL_CORP_GRP, COL_CUST_NAME,
    COL_FUEL_SEG, COL_LUBE_CAT,
    COL_GRS_CY, COL_GRS_SPLY, COL_VOL_CY, COL_VOL_SPLY,
    COL_DISC_CY, COL_DISC_SPLY,
    COL_PMGN_CY, COL_PMGN_SPLY, COL_NMGN_CY, COL_NMGN_SPLY,
    COL_IS_RETAIL,
    LUBE_CATEGORIES,
)
from pso.analyze import _agg, _vol_share, _pareto, _pct_chg


def run_lubes(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Full lubricants analysis on Retail Business rows.
    Returns dict of named DataFrames.
    """
    lubes = df[df[COL_IS_RETAIL] & (df[COL_FUEL_SEG] == "Lubricants")].copy()

    tables: dict[str, pd.DataFrame] = {}

    # 1. Category trend — the headline view
    tables["lube_category_trend"]  = _lube_category_trend(lubes)

    # 2. Margin decomposition per category
    tables["lube_margin_decomp"]   = _margin_decomp(lubes)

    # 3. Region × Category matrix (volume)
    tables["lube_region_category"] = _region_category_matrix(lubes)

    # 4. Top 30 cities × Category (volume)
    tables["lube_city_category"]   = _city_category(lubes, top_n=30)

    # 5. Top 50 lube customers
    tables["lube_top_customers"]   = _top_customers(lubes, n=50)

    # 6. Declining markets
    tables["lube_declining_mkts"]  = _volume_movers(lubes, direction="decline")

    # 7. Growing markets
    tables["lube_growing_mkts"]    = _volume_movers(lubes, direction="growth")

    # 8. Discount leakage by city
    tables["lube_discount_city"]   = _discount_by_city(lubes)

    # 9. NetMargin/Ltr pockets — category × region
    tables["lube_margin_pockets"]  = _margin_per_ltr_matrix(lubes)

    # 10. LOW GRADE specific deep dive
    low_grade = lubes[lubes[COL_LUBE_CAT] == "LOW GRADE"]
    if len(low_grade):
        tables["low_grade_detail"]  = _category_detail(low_grade, "LOW GRADE")

    # 11. DEO deep dive
    deo = lubes[lubes[COL_LUBE_CAT] == "DEO"]
    if len(deo):
        tables["deo_detail"] = _category_detail(deo, "DEO")

    # 12. PCMO deep dive
    pcmo = lubes[lubes[COL_LUBE_CAT] == "PCMO"]
    if len(pcmo):
        tables["pcmo_detail"] = _category_detail(pcmo, "PCMO")

    # 13. MCO deep dive
    mco = lubes[lubes[COL_LUBE_CAT] == "MCO"]
    if len(mco):
        tables["mco_detail"] = _category_detail(mco, "MCO")

    # 14. Corporate group / customer segment analysis
    tables["lube_corp_segments"]   = _corp_segment_analysis(lubes)

    return tables


# ── builders ──────────────────────────────────────────────────────────────────

def _lube_category_trend(lubes: pd.DataFrame) -> pd.DataFrame:
    rows = []
    # Include total row
    groups = [(None, lubes)] + [(cat, lubes[lubes[COL_LUBE_CAT] == cat])
                                 for cat in LUBE_CATEGORIES]
    for cat, grp in groups:
        label = cat if cat else "ALL LUBES"
        vol_cy = grp[COL_VOL_CY].sum()
        vol_sply = grp[COL_VOL_SPLY].sum()
        grs_cy = grp[COL_GRS_CY].sum()
        grs_sply = grp[COL_GRS_SPLY].sum()
        pm_cy  = grp[COL_PMGN_CY].sum()
        pm_sply  = grp[COL_PMGN_SPLY].sum()
        disc_cy = grp[COL_DISC_CY].sum()
        disc_sply = grp[COL_DISC_SPLY].sum()
        nm_cy  = grp[COL_NMGN_CY].sum()
        nm_sply  = grp[COL_NMGN_SPLY].sum()

        rows.append({
            "Category":          label,
            "Rows_CY":           len(grp),
            "Vol_CY_ML":         round(vol_cy / 1e6, 3),
            "Vol_SPLY_ML":       round(vol_sply / 1e6, 3),
            "Vol_Chg_Pct":       round(_pct_chg(vol_cy, vol_sply), 1) if vol_sply else np.nan,
            "GRS_CY_B":          round(grs_cy / 1e9, 3),
            "GRS_SPLY_B":        round(grs_sply / 1e9, 3),
            "GRS_Chg_Pct":       round(_pct_chg(grs_cy, grs_sply), 1) if grs_sply else np.nan,
            "PrimaryMgn_CY_B":   round(pm_cy / 1e9, 3),
            "PrimaryMgn_SPLY_B": round(pm_sply / 1e9, 3),
            "Discount_CY_B":     round(disc_cy / 1e9, 3),
            "Discount_SPLY_B":   round(disc_sply / 1e9, 3),
            "NetMargin_CY_B":    round(nm_cy / 1e9, 3),
            "NetMargin_SPLY_B":  round(nm_sply / 1e9, 3),
            "NMgn_Chg_Pct":      round(_pct_chg(nm_cy, nm_sply), 1) if nm_sply else np.nan,
            "NMgn_per_Ltr_CY":   round(nm_cy / vol_cy, 3) if vol_cy else np.nan,
            "NMgn_per_Ltr_SPLY": round(nm_sply / vol_sply, 3) if vol_sply else np.nan,
        })

    return pd.DataFrame(rows)


def _margin_decomp(lubes: pd.DataFrame) -> pd.DataFrame:
    """
    Per category: Primary Margin → less Discount → Net Margin (all per litre).
    Flags compression and leakage.
    """
    rows = []
    for cat in LUBE_CATEGORIES:
        grp = lubes[lubes[COL_LUBE_CAT] == cat]
        if grp.empty:
            continue
        vol_cy  = grp[COL_VOL_CY].sum()
        vol_sply  = grp[COL_VOL_SPLY].sum()
        pm_cy   = grp[COL_PMGN_CY].sum()
        pm_sply   = grp[COL_PMGN_SPLY].sum()
        disc_cy = grp[COL_DISC_CY].sum()
        disc_sply = grp[COL_DISC_SPLY].sum()
        nm_cy   = grp[COL_NMGN_CY].sum()
        nm_sply   = grp[COL_NMGN_SPLY].sum()

        def per_l(val, vol):
            return round(val / vol, 3) if vol else np.nan

        rows.append({
            "Category":               cat,
            "Vol_CY_ML":              round(vol_cy / 1e6, 3),
            "PMgn_per_Ltr_CY":        per_l(pm_cy, vol_cy),
            "PMgn_per_Ltr_SPLY":      per_l(pm_sply, vol_sply),
            "PMgn_per_Ltr_Chg":       round(per_l(pm_cy, vol_cy) - per_l(pm_sply, vol_sply), 3)
                                       if (vol_cy and vol_sply) else np.nan,
            "Disc_per_Ltr_CY":        per_l(disc_cy, vol_cy),
            "Disc_per_Ltr_SPLY":      per_l(disc_sply, vol_sply),
            "Disc_per_Ltr_Chg":       round(per_l(disc_cy, vol_cy) - per_l(disc_sply, vol_sply), 3)
                                       if (vol_cy and vol_sply) else np.nan,
            "NMgn_per_Ltr_CY":        per_l(nm_cy, vol_cy),
            "NMgn_per_Ltr_SPLY":      per_l(nm_sply, vol_sply),
            "NMgn_per_Ltr_Chg":       round(per_l(nm_cy, vol_cy) - per_l(nm_sply, vol_sply), 3)
                                       if (vol_cy and vol_sply) else np.nan,
            "Disc_Increased":         (per_l(disc_cy, vol_cy) or 0) > (per_l(disc_sply, vol_sply) or 0),
            "Margin_Compressed":      (per_l(nm_cy, vol_cy) or 0) < (per_l(nm_sply, vol_sply) or 0),
            "Disc_as_Pct_of_PMgn":    round(disc_cy / pm_cy * 100, 1) if pm_cy else np.nan,
        })
    return pd.DataFrame(rows)


def _region_category_matrix(lubes: pd.DataFrame) -> pd.DataFrame:
    """Pivot: rows=Region, cols=LubeCategory, values=Vol_CY_ML."""
    t = (
        lubes.groupby([COL_REGION, COL_LUBE_CAT], dropna=False)
        .agg(Vol_CY_ML=(COL_VOL_CY, lambda x: round(x.sum() / 1e6, 3)),
             NMgn_CY_B=(COL_NMGN_CY, lambda x: round(x.sum() / 1e9, 3)))
        .reset_index()
    )
    # Volume pivot
    vol_piv = t.pivot_table(
        index=COL_REGION, columns=COL_LUBE_CAT,
        values="Vol_CY_ML", aggfunc="sum", fill_value=0
    )
    vol_piv.columns = [f"Vol_{c}_ML" for c in vol_piv.columns]
    # Margin pivot
    mgn_piv = t.pivot_table(
        index=COL_REGION, columns=COL_LUBE_CAT,
        values="NMgn_CY_B", aggfunc="sum", fill_value=0
    )
    mgn_piv.columns = [f"NMgn_{c}_B" for c in mgn_piv.columns]
    result = pd.concat([vol_piv, mgn_piv], axis=1).reset_index()
    result.columns.name = None
    return result


def _city_category(lubes: pd.DataFrame, top_n: int = 30) -> pd.DataFrame:
    """Top cities by total lube volume, with category breakdown."""
    city_total = (
        lubes.groupby(COL_CITY_NORM)[COL_VOL_CY]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
    )
    top_cities = city_total.index.tolist()
    subset = lubes[lubes[COL_CITY_NORM].isin(top_cities)]

    piv = subset.pivot_table(
        index=[COL_CITY_NORM, COL_REGION],
        columns=COL_LUBE_CAT,
        values=COL_VOL_CY,
        aggfunc="sum",
        fill_value=0,
    )
    piv["Total_Vol_CY"] = piv.sum(axis=1)
    piv.columns = [f"Vol_{c}_L" if c != "Total_Vol_CY" else c for c in piv.columns]
    result = piv.reset_index().sort_values("Total_Vol_CY", ascending=False)
    result["Total_Vol_CY_ML"] = (result["Total_Vol_CY"] / 1e6).round(3)
    result.columns.name = None
    return result


def _top_customers(lubes: pd.DataFrame, n: int = 50) -> pd.DataFrame:
    t = (
        lubes.groupby([COL_CUST_NAME, COL_REGION, COL_LUBE_CAT], dropna=False)
        .agg(
            Vol_CY_ML=(COL_VOL_CY, lambda x: round(x.sum() / 1e6, 4)),
            Vol_SPLY_ML=(COL_VOL_SPLY, lambda x: round(x.sum() / 1e6, 4)),
            NMgn_CY_B=(COL_NMGN_CY, lambda x: round(x.sum() / 1e9, 4)),
        )
        .reset_index()
    )
    # Summarise at customer level
    cust_total = (
        t.groupby([COL_CUST_NAME, COL_REGION], dropna=False)
        .agg(
            Total_Vol_CY_ML=("Vol_CY_ML", "sum"),
            Total_Vol_SPLY_ML=("Vol_SPLY_ML", "sum"),
            Total_NMgn_CY_B=("NMgn_CY_B", "sum"),
        )
        .reset_index()
        .sort_values("Total_Vol_CY_ML", ascending=False)
        .head(n)
        .reset_index(drop=True)
    )
    cust_total["Vol_Chg_Pct"] = _pct_chg(
        cust_total["Total_Vol_CY_ML"], cust_total["Total_Vol_SPLY_ML"]
    ).round(1)
    cust_total["Rank"] = cust_total.index + 1
    return cust_total[["Rank", COL_CUST_NAME, COL_REGION,
                         "Total_Vol_CY_ML", "Total_Vol_SPLY_ML",
                         "Vol_Chg_Pct", "Total_NMgn_CY_B"]]


def _volume_movers(lubes: pd.DataFrame, direction: str = "decline") -> pd.DataFrame:
    """Cities where lube volume declined/grew most vs SPLY (absolute change)."""
    city = (
        lubes.groupby([COL_CITY_NORM, COL_REGION])
        .agg(Vol_CY_ML=(COL_VOL_CY, lambda x: x.sum() / 1e6),
             Vol_SPLY_ML=(COL_VOL_SPLY, lambda x: x.sum() / 1e6))
        .reset_index()
    )
    city["Vol_Delta_ML"] = city["Vol_CY_ML"] - city["Vol_SPLY_ML"]
    city["Vol_Chg_Pct"] = _pct_chg(city["Vol_CY_ML"], city["Vol_SPLY_ML"]).round(1)
    if direction == "decline":
        return city[city["Vol_Delta_ML"] < 0].sort_values("Vol_Delta_ML").head(30).reset_index(drop=True)
    else:
        return city[city["Vol_Delta_ML"] > 0].sort_values("Vol_Delta_ML", ascending=False).head(30).reset_index(drop=True)


def _discount_by_city(lubes: pd.DataFrame) -> pd.DataFrame:
    city = (
        lubes.groupby([COL_CITY_NORM, COL_REGION])
        .agg(
            Vol_CY_ML =(COL_VOL_CY,  lambda x: x.sum() / 1e6),
            Disc_CY_B =(COL_DISC_CY, lambda x: x.sum() / 1e9),
            PMgn_CY_B =(COL_PMGN_CY, lambda x: x.sum() / 1e9),
            NMgn_CY_B =(COL_NMGN_CY, lambda x: x.sum() / 1e9),
        )
        .reset_index()
    )
    city["Disc_per_Ltr_CY"] = np.where(
        city["Vol_CY_ML"] > 0,
        city["Disc_CY_B"] * 1e9 / (city["Vol_CY_ML"] * 1e6),
        np.nan,
    ).round(3)
    city["NMgn_per_Ltr_CY"] = np.where(
        city["Vol_CY_ML"] > 0,
        city["NMgn_CY_B"] * 1e9 / (city["Vol_CY_ML"] * 1e6),
        np.nan,
    ).round(3)
    return city.sort_values("Disc_per_Ltr_CY", ascending=False).head(40).reset_index(drop=True)


def _margin_per_ltr_matrix(lubes: pd.DataFrame) -> pd.DataFrame:
    """Avg Net Margin/Ltr by Region × Category — reveals value pockets."""
    rows = []
    for region, rgrp in lubes.groupby(COL_REGION):
        for cat in LUBE_CATEGORIES:
            cgrp = rgrp[rgrp[COL_LUBE_CAT] == cat]
            if cgrp.empty:
                continue
            vol = cgrp[COL_VOL_CY].sum()
            nm  = cgrp[COL_NMGN_CY].sum()
            disc = cgrp[COL_DISC_CY].sum()
            pm  = cgrp[COL_PMGN_CY].sum()
            rows.append({
                COL_REGION:            region,
                "LubeCategory":        cat,
                "Vol_CY_ML":           round(vol / 1e6, 3),
                "PMgn_per_Ltr":        round(pm / vol, 3) if vol else np.nan,
                "Disc_per_Ltr":        round(disc / vol, 3) if vol else np.nan,
                "NMgn_per_Ltr":        round(nm / vol, 3) if vol else np.nan,
                "Disc_as_Pct_PMgn":    round(disc / pm * 100, 1) if pm else np.nan,
            })
    return pd.DataFrame(rows).sort_values([COL_REGION, "NMgn_per_Ltr"], ascending=[True, False])


def _category_detail(grp: pd.DataFrame, cat_name: str) -> pd.DataFrame:
    """Region + top-20-city breakdown for a single lube category."""
    by_region = _agg(grp, [COL_REGION]).pipe(_vol_share)
    by_region.insert(0, "Breakdown", "By Region")

    by_city = (
        _agg(grp, [COL_REGION, COL_CITY_NORM])
        .sort_values("Vol_CY_ML", ascending=False)
        .head(20)
        .reset_index(drop=True)
    )
    total_vol = grp[COL_VOL_CY].sum() / 1e6
    by_city["Vol_Share_Pct"] = (by_city["Vol_CY_ML"] / total_vol * 100).round(2)
    by_city.insert(0, "Breakdown", "By City (top 20)")

    by_corp = (
        grp.groupby([COL_CORP_GRP, COL_REGION], dropna=False)
        .agg(Vol_CY_ML=(COL_VOL_CY, lambda x: round(x.sum() / 1e6, 4)),
             NMgn_CY_B=(COL_NMGN_CY, lambda x: round(x.sum() / 1e9, 4)))
        .reset_index()
        .sort_values("Vol_CY_ML", ascending=False)
        .head(15)
        .reset_index(drop=True)
    )
    by_corp.insert(0, "Breakdown", "By Corp Group (top 15)")

    return pd.concat([by_region, by_city, by_corp], ignore_index=True)


def _corp_segment_analysis(lubes: pd.DataFrame) -> pd.DataFrame:
    """Which corporate segments (cartage, IC, etc.) buy lubes and what types."""
    # Named corporate groups only (not blank, not numeric-only)
    named = lubes[lubes[COL_CORP_GRP].str.len() > 3 & ~lubes[COL_CORP_GRP].str.isnumeric()]

    t = (
        named.groupby([COL_CORP_GRP, COL_LUBE_CAT], dropna=False)
        .agg(Vol_CY_ML=(COL_VOL_CY, lambda x: round(x.sum() / 1e6, 4)))
        .reset_index()
    )
    total = t.groupby(COL_CORP_GRP)["Vol_CY_ML"].sum().sort_values(ascending=False).head(20)
    top_groups = total.index.tolist()
    result = t[t[COL_CORP_GRP].isin(top_groups)].copy()
    result = result.sort_values([COL_CORP_GRP, "Vol_CY_ML"], ascending=[True, False])
    return result.reset_index(drop=True)
