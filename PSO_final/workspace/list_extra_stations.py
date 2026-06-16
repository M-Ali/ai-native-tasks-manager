"""List all 92 web-resolved stations with names and current district assignments."""
import sys, json, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'src')

from pso import ingest

df, _ = ingest.load("data/input/Working File Retail Fuels Data.xlsx")
retail = df[df["IsRetail"] & ~df["IsInternational"]].copy()
kar = retail[retail["CityNorm"] == "Karachi"].copy()
kar["Customer Number"] = kar["Customer Number"].astype(str).str.strip()

kar_stns = (kar.groupby("Customer Number", as_index=False)
            .agg(name=("Name 1","first"),
                 ml_cy=("SalesLtr_CY", lambda x: x.sum()/1e6)))

pumps = pd.read_csv("karachi_pumps.csv", dtype={"Code": str})
pumps["Code"] = pumps["Code"].str.strip().str.zfill(10)
pumps_kar = pumps[pumps["Division"].str.strip() == "Karachi"].copy()
code_to_dist = pumps_kar.set_index("Code")["City_District_Area"].to_dict()

with open("workspace/extra_district_map.json") as f:
    extra_raw = json.load(f)
extra_map = {k: v for k, v in extra_raw.items() if not k.startswith("_")}

# Only the 92 extra-map stations (not in CSV)
extra_mask = (kar_stns["Customer Number"].isin(extra_map) &
              ~kar_stns["Customer Number"].isin(code_to_dist))
extra_stns = kar_stns[extra_mask].copy()
extra_stns["District"] = extra_stns["Customer Number"].map(extra_map)
extra_stns = extra_stns.sort_values("ml_cy", ascending=False).reset_index(drop=True)

print(f"{'#':<4} {'Code':<12} {'CY(ML)':>7}  {'District':<22}  Name")
print("-" * 100)
for i, (_, r) in enumerate(extra_stns.iterrows(), 1):
    print(f"{i:<4} {r['Customer Number']:<12} {r['ml_cy']:>7.2f}  {r['District']:<22}  {r['name']}")
print(f"\nTotal: {len(extra_stns)} stations, {extra_stns['ml_cy'].sum():.1f}ML")
