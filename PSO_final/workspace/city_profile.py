"""
PSO — Top 10 City Profiles
Per city: station count, volume, GRS, net margin, performing vs non-performing stations.

Performing   : Vol_CY > 0 AND NetMargin_CY > 0
Non-performing: Vol_CY == 0 (inactive) OR NetMargin_CY <= 0
"""

import sys, os
sys.path.insert(0, 'src')
os.environ.setdefault('PYTHONIOENCODING', 'utf-8')

import pandas as pd
from pso import ingest

DATA_FILE = "data/input/Working File Retail Fuels Data.xlsx"
df, _ = ingest.load(DATA_FILE)

# ── Retail fuel rows only (exclude lubes for station count clarity) ────────────
retail = df[df["IsRetail"] & ~df["IsInternational"]].copy()

# ── Aggregate per Customer (station) × City ────────────────────────────────────
# A customer can appear multiple times (different products/segments)
# Collapse to one row per station per city
station_city = (
    retail
    .groupby(["CityNorm", "Customer Number"], as_index=False)
    .agg(
        GRS_CY     = ("SalesGRS_CY",    "sum"),
        Vol_CY_ML  = ("SalesLtr_CY",    lambda x: x.sum() / 1_000_000),
        NMgn_CY    = ("NetMargin_CY",   "sum"),
        GRS_LY     = ("SalesGRS_SPLY",    "sum"),
        Vol_LY_ML  = ("SalesLtr_SPLY",    lambda x: x.sum() / 1_000_000),
    )
)

# Classify each station
station_city["performing"]     = (station_city["Vol_CY_ML"] > 0) & (station_city["NMgn_CY"] > 0)
station_city["non_performing"] = ~station_city["performing"]
station_city["inactive"]       = station_city["Vol_CY_ML"] == 0

# ── City-level aggregation ─────────────────────────────────────────────────────
city_profile = (
    station_city
    .groupby("CityNorm", as_index=False)
    .agg(
        Total_Stations    = ("Customer Number",  "count"),
        Performing        = ("performing",        "sum"),
        Non_Performing    = ("non_performing",    "sum"),
        Inactive          = ("inactive",          "sum"),
        GRS_CY_PKR_M      = ("GRS_CY",           lambda x: x.sum() / 1_000_000),
        Vol_CY_ML         = ("Vol_CY_ML",         "sum"),
        NMgn_CY_PKR_M     = ("NMgn_CY",          lambda x: x.sum() / 1_000_000),
        GRS_LY_PKR_M      = ("GRS_LY",           lambda x: x.sum() / 1_000_000),
        Vol_LY_ML         = ("Vol_LY_ML",         "sum"),
    )
)

# Growth columns
city_profile["Vol_Chg_pct"]  = ((city_profile["Vol_CY_ML"]  - city_profile["Vol_LY_ML"])
                                 / city_profile["Vol_LY_ML"].replace(0, float("nan")) * 100).round(1)
city_profile["GRS_Chg_pct"]  = ((city_profile["GRS_CY_PKR_M"] - city_profile["GRS_LY_PKR_M"])
                                 / city_profile["GRS_LY_PKR_M"].replace(0, float("nan")) * 100).round(1)

# Total for share calc
total_grs = city_profile["GRS_CY_PKR_M"].sum()
total_vol = city_profile["Vol_CY_ML"].sum()
city_profile["GRS_Share_pct"] = (city_profile["GRS_CY_PKR_M"] / total_grs * 100).round(1)
city_profile["Vol_Share_pct"] = (city_profile["Vol_CY_ML"]    / total_vol * 100).round(1)
city_profile["NMgn_per_Ltr"]  = (city_profile["NMgn_CY_PKR_M"] * 1_000_000
                                  / (city_profile["Vol_CY_ML"] * 1_000_000)
                                 ).replace([float("inf"), float("-inf")], 0).round(2)
city_profile["Performing_pct"]= (city_profile["Performing"]
                                  / city_profile["Total_Stations"] * 100).round(0).astype(int)

# Top 10 by GRS (value)
top10 = city_profile.nlargest(10, "GRS_CY_PKR_M").reset_index(drop=True)
top10.index += 1

display_cols = [
    "CityNorm", "Total_Stations", "Performing", "Non_Performing", "Inactive",
    "Performing_pct",
    "GRS_CY_PKR_M", "GRS_Share_pct", "GRS_Chg_pct",
    "Vol_CY_ML",    "Vol_Share_pct", "Vol_Chg_pct",
    "NMgn_CY_PKR_M", "NMgn_per_Ltr",
]
top10_disp = top10[display_cols].copy()
top10_disp.columns = [
    "City", "Stations", "Performing", "Non-Perf", "Inactive", "Perf %",
    "GRS CY (PKR M)", "GRS Share%", "GRS Chg%",
    "Vol CY (ML)", "Vol Share%", "Vol Chg%",
    "NMgn CY (PKR M)", "NMgn/ltr",
]

pd.set_option("display.width", 200)
pd.set_option("display.max_columns", 20)
pd.set_option("display.float_format", "{:,.1f}".format)
print("\n=== TOP 10 CITIES BY VALUE (GRS) ===")
print(top10_disp.to_string())

# ── Cumulative share check ─────────────────────────────────────────────────────
print(f"\nTop 10 cumulative GRS share: {top10['GRS_Share_pct'].sum():.1f}%")
print(f"Top 10 cumulative Vol share: {top10['Vol_Share_pct'].sum():.1f}%")
print(f"Top 10 total stations: {top10['Total_Stations'].sum()}")
print(f"Top 10 total performing: {top10['Performing'].sum()}")
print(f"Top 10 total non-performing: {top10['Non_Performing'].sum()}")
print(f"Avg performing rate: {(top10['Performing'] / top10['Total_Stations'] * 100).mean():.0f}%")
