"""Premium Fuel Agent — R95 (premium) vs PMG (regular) petrol deep-dive.

PSO already sells a premium petrol variant (R95) alongside regular petrol (PMG) —
both classified under FuelSegment == "Petrol". Rather than speculate about a
hypothetical new high-end fuel, this agent uses R95's existing footprint as the
proof-of-concept: where R95 already over-indexes proves real demand exists; where
an active PMG station carries zero R95 is the direct "where to launch/expand"
whitespace list. CY vs SPLY throughout — never full-year LY (see pso.period).
"""

from __future__ import annotations

import pandas as pd
import numpy as np

from pso.config import (
    COL_PRODUCT, COL_REGION, COL_CITY_NORM, COL_CORP_GRP, COL_CUST_NAME, COL_CUST_NUM,
    COL_FUEL_SEG, COL_IS_RETAIL,
    COL_VOL_CY, COL_VOL_SPLY,
    COL_DISC_CY, COL_DISC_SPLY, COL_PMGN_CY, COL_PMGN_SPLY, COL_NMGN_CY, COL_NMGN_SPLY,
)
from pso.analyze import _agg, _vol_share, _pct_chg

PREMIUM_PRODUCT = "R95"
REGULAR_PRODUCT = "PMG"


def run_premium_fuel(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Full R95-vs-PMG analysis on Retail Business + Petrol rows. Returns dict of
    named DataFrames."""
    petrol = df[df[COL_IS_RETAIL] & (df[COL_FUEL_SEG] == "Petrol")].copy()
    petrol = petrol[petrol[COL_PRODUCT].isin([PREMIUM_PRODUCT, REGULAR_PRODUCT])]

    tables: dict[str, pd.DataFrame] = {}

    tables["premium_product_trend"]   = _agg(petrol, [COL_PRODUCT]).pipe(_vol_share)
    tables["premium_margin_decomp"]   = _margin_decomp(petrol)
    tables["premium_by_region"]       = _region_mix(petrol)
    tables["premium_by_city"]         = _city_mix(petrol)

    station_mix = _station_mix(petrol)
    tables["premium_station_mix"]        = station_mix
    tables["premium_whitespace_stations"] = _whitespace(station_mix)

    growing, declining = _volume_movers(petrol)
    tables["premium_growing_markets"]   = growing
    tables["premium_declining_markets"] = declining

    tables["premium_customer_segments"] = _customer_segments(petrol)

    return tables


# ── builders ──────────────────────────────────────────────────────────────────

def _margin_decomp(petrol: pd.DataFrame) -> pd.DataFrame:
    """Primary Margin > Discount > Net Margin per litre, PMG vs R95, plus the
    R95-minus-PMG uplift row (the unit-economics case for pushing premium fuel)."""
    rows = []
    for prod in [REGULAR_PRODUCT, PREMIUM_PRODUCT]:
        grp = petrol[petrol[COL_PRODUCT] == prod]
        if grp.empty:
            continue
        vol_cy   = grp[COL_VOL_CY].sum()
        vol_sply = grp[COL_VOL_SPLY].sum()
        pm_cy    = grp[COL_PMGN_CY].sum()
        pm_sply  = grp[COL_PMGN_SPLY].sum()
        disc_cy   = grp[COL_DISC_CY].sum()
        disc_sply = grp[COL_DISC_SPLY].sum()
        nm_cy    = grp[COL_NMGN_CY].sum()
        nm_sply  = grp[COL_NMGN_SPLY].sum()

        def per_l(val, vol):
            return round(val / vol, 3) if vol else np.nan

        nm_l_cy, nm_l_sply = per_l(nm_cy, vol_cy), per_l(nm_sply, vol_sply)
        rows.append({
            "Product":              prod,
            "Vol_CY_ML":            round(vol_cy / 1e6, 3),
            "Vol_SPLY_ML":          round(vol_sply / 1e6, 3),
            "Vol_Chg_Pct":          round(_pct_chg(vol_cy, vol_sply), 1) if vol_sply else np.nan,
            "PMgn_per_Ltr_CY":      per_l(pm_cy, vol_cy),
            "PMgn_per_Ltr_SPLY":    per_l(pm_sply, vol_sply),
            "Disc_per_Ltr_CY":      per_l(disc_cy, vol_cy),
            "Disc_per_Ltr_SPLY":    per_l(disc_sply, vol_sply),
            "NMgn_per_Ltr_CY":      nm_l_cy,
            "NMgn_per_Ltr_SPLY":    nm_l_sply,
            "NMgn_per_Ltr_Chg_Pct": round(_pct_chg(nm_l_cy, nm_l_sply), 1)
                                      if (not np.isnan(nm_l_cy) and not np.isnan(nm_l_sply) and nm_l_sply != 0)
                                      else np.nan,
        })

    result = pd.DataFrame(rows)
    if {REGULAR_PRODUCT, PREMIUM_PRODUCT}.issubset(set(result.get("Product", []))):
        r95 = result[result["Product"] == PREMIUM_PRODUCT].iloc[0]
        pmg = result[result["Product"] == REGULAR_PRODUCT].iloc[0]
        uplift = pd.DataFrame([{
            "Product":              f"{PREMIUM_PRODUCT} minus {REGULAR_PRODUCT} (per-litre uplift)",
            "Vol_CY_ML":            np.nan,
            "Vol_SPLY_ML":          np.nan,
            "Vol_Chg_Pct":          np.nan,
            "PMgn_per_Ltr_CY":      round(r95["PMgn_per_Ltr_CY"] - pmg["PMgn_per_Ltr_CY"], 3),
            "PMgn_per_Ltr_SPLY":    np.nan,
            "Disc_per_Ltr_CY":      round(r95["Disc_per_Ltr_CY"] - pmg["Disc_per_Ltr_CY"], 3),
            "Disc_per_Ltr_SPLY":    np.nan,
            "NMgn_per_Ltr_CY":      round(r95["NMgn_per_Ltr_CY"] - pmg["NMgn_per_Ltr_CY"], 3),
            "NMgn_per_Ltr_SPLY":    np.nan,
            "NMgn_per_Ltr_Chg_Pct": np.nan,
        }])
        result = pd.concat([result, uplift], ignore_index=True)
    return result


def _product_pivot(petrol: pd.DataFrame, by: str | list[str]) -> pd.DataFrame:
    """Shared helper: volume pivot of PMG/R95 by an arbitrary grouping key(s),
    with R95 share of the group's petrol volume."""
    t = (
        petrol.groupby(([by] if isinstance(by, str) else by) + [COL_PRODUCT], dropna=False)[COL_VOL_CY]
        .sum().unstack(fill_value=0)
    )
    t.columns.name = None
    for prod in (REGULAR_PRODUCT, PREMIUM_PRODUCT):
        if prod not in t.columns:
            t[prod] = 0
    t["PMG_Vol_CY_ML"]   = t[REGULAR_PRODUCT] / 1e6
    t["R95_Vol_CY_ML"]   = t[PREMIUM_PRODUCT] / 1e6
    t["Total_Vol_CY_ML"] = t["PMG_Vol_CY_ML"] + t["R95_Vol_CY_ML"]
    t["R95_Share_Pct"] = np.where(
        t["Total_Vol_CY_ML"] > 0, t["R95_Vol_CY_ML"] / t["Total_Vol_CY_ML"] * 100, 0
    ).round(2)
    return t[["PMG_Vol_CY_ML", "R95_Vol_CY_ML", "Total_Vol_CY_ML", "R95_Share_Pct"]].reset_index()


def _region_mix(petrol: pd.DataFrame) -> pd.DataFrame:
    """R95 share of regional petrol volume, ranked — where premium already sells."""
    result = _product_pivot(petrol, COL_REGION)
    n_stns = (
        petrol[petrol[COL_PRODUCT] == PREMIUM_PRODUCT]
        .groupby(COL_REGION)[COL_CUST_NUM].nunique()
        .rename("R95_Stations")
    )
    result = result.merge(n_stns, on=COL_REGION, how="left")
    result["R95_Stations"] = result["R95_Stations"].fillna(0).astype(int)
    return result.sort_values("R95_Share_Pct", ascending=False).reset_index(drop=True)


def _city_mix(petrol: pd.DataFrame) -> pd.DataFrame:
    """R95 penetration % by city, ranked — proof-of-demand cities."""
    result = _product_pivot(petrol, [COL_REGION, COL_CITY_NORM])
    result = result[result["Total_Vol_CY_ML"] > 0]
    return result.sort_values("R95_Share_Pct", ascending=False).reset_index(drop=True)


def _station_mix(petrol: pd.DataFrame) -> pd.DataFrame:
    """Station-level R95 vs PMG volume + R95 % of the station's own petrol
    volume, ranked by R95 volume — "which stations already sell more expensive
    fuel"."""
    t = (
        petrol.groupby([COL_CUST_NUM, COL_CUST_NAME, COL_CITY_NORM, COL_REGION, COL_PRODUCT], dropna=False)[COL_VOL_CY]
        .sum().unstack(fill_value=0)
    )
    t.columns.name = None
    t = t.reset_index()
    for prod in (REGULAR_PRODUCT, PREMIUM_PRODUCT):
        if prod not in t.columns:
            t[prod] = 0
    t["PMG_Vol_CY_ML"]   = t[REGULAR_PRODUCT] / 1e6
    t["R95_Vol_CY_ML"]   = t[PREMIUM_PRODUCT] / 1e6
    t["Total_Vol_CY_ML"] = t["PMG_Vol_CY_ML"] + t["R95_Vol_CY_ML"]
    t["R95_Share_Pct"] = np.where(
        t["Total_Vol_CY_ML"] > 0, t["R95_Vol_CY_ML"] / t["Total_Vol_CY_ML"] * 100, 0
    ).round(2)
    result = t[[COL_CUST_NUM, COL_CUST_NAME, COL_CITY_NORM, COL_REGION,
                "PMG_Vol_CY_ML", "R95_Vol_CY_ML", "Total_Vol_CY_ML", "R95_Share_Pct"]]
    return result.sort_values("R95_Vol_CY_ML", ascending=False).reset_index(drop=True)


def _whitespace(station_mix: pd.DataFrame, min_pmg_ml: float = 0.01) -> pd.DataFrame:
    """Active PMG stations carrying zero R95 — ranked by PMG volume (size of
    the opportunity). Direct "where to launch" candidate list."""
    ws = station_mix[
        (station_mix["R95_Vol_CY_ML"] <= 0) & (station_mix["PMG_Vol_CY_ML"] >= min_pmg_ml)
    ].copy()
    return ws.sort_values("PMG_Vol_CY_ML", ascending=False).reset_index(drop=True)


def _volume_movers(petrol: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cities where R95 volume grew/declined most vs SPLY (absolute change) —
    momentum signal for where premium demand is building or eroding."""
    r95 = petrol[petrol[COL_PRODUCT] == PREMIUM_PRODUCT]
    city = (
        r95.groupby([COL_CITY_NORM, COL_REGION])
        .agg(Vol_CY_ML=(COL_VOL_CY, lambda x: x.sum() / 1e6),
             Vol_SPLY_ML=(COL_VOL_SPLY, lambda x: x.sum() / 1e6))
        .reset_index()
    )
    city["Vol_Delta_ML"] = city["Vol_CY_ML"] - city["Vol_SPLY_ML"]
    city["Vol_Chg_Pct"] = _pct_chg(city["Vol_CY_ML"], city["Vol_SPLY_ML"]).round(1)
    growing   = city[city["Vol_Delta_ML"] > 0].sort_values("Vol_Delta_ML", ascending=False).head(30).reset_index(drop=True)
    declining = city[city["Vol_Delta_ML"] < 0].sort_values("Vol_Delta_ML").head(30).reset_index(drop=True)
    return growing, declining


def _customer_segments(petrol: pd.DataFrame) -> pd.DataFrame:
    """Corporate-group / customer mix buying R95 — fleet vs retail-walk-in proxy."""
    r95 = petrol[petrol[COL_PRODUCT] == PREMIUM_PRODUCT]
    named = r95[r95[COL_CORP_GRP].str.len() > 3 & ~r95[COL_CORP_GRP].str.isnumeric()]
    t = (
        named.groupby(COL_CORP_GRP, dropna=False)
        .agg(R95_Vol_CY_ML=(COL_VOL_CY, lambda x: round(x.sum() / 1e6, 4)),
             R95_Stations=(COL_CUST_NUM, "nunique"))
        .reset_index()
        .sort_values("R95_Vol_CY_ML", ascending=False)
        .head(20)
        .reset_index(drop=True)
    )
    return t
