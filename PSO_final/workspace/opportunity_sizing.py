"""
PSO OMC — City Opportunity Sizing
Two opportunity types per city:
  A. Lube mix opportunity  (LOW GRADE -> DEO/PCMO uplift)
  B. Fuel discount normalisation (vs South benchmark)
"""

import os, sys
sys.path.insert(0, 'src')

import pandas as pd
from pso import ingest, analyze, lubes_analyze

DATA_FILE = "data/input/Working File Retail Fuels Data.xlsx"

print("Loading pipeline tables ...")
df, _           = ingest.load(DATA_FILE)
analysis_tables = analyze.run_all(df)
lubes_tables    = lubes_analyze.run_lubes(df)

# ── Inspect table structures (debug) ─────────────────────────────────────────
def cols(tbl_name, tables):
    t = tables.get(tbl_name, pd.DataFrame())
    if t.empty:
        return None
    return list(t.columns)

# ── A. LUBE MIX OPPORTUNITY ───────────────────────────────────────────────────
print("\n== A. LUBE MIX OPPORTUNITY (top 15 cities) ==")

# lube_city_category is wide: CityNorm, Region, Vol_DEO_L, Vol_LOW GRADE_L, Vol_PCMO_L, ...
lcc = lubes_tables.get('lube_city_category', pd.DataFrame())
if not lcc.empty:
    print(f"  Columns: {list(lcc.columns)}")

    city_col = 'CityNorm'
    lowg_col = next((c for c in lcc.columns if 'LOW GRADE' in c and 'Vol' in c), None)
    deo_col  = next((c for c in lcc.columns if 'DEO' in c and 'Vol' in c), None)
    pcmo_col = next((c for c in lcc.columns if 'PCMO' in c and 'Vol' in c), None)
    mco_col  = next((c for c in lcc.columns if 'MCO' in c and 'Vol' in c and 'PCMO' not in c), None)
    tot_col  = next((c for c in lcc.columns if 'Total_Vol' in c and 'ML' in c), None)

    print(f"  LOW GRADE={lowg_col}  DEO={deo_col}  PCMO={pcmo_col}  Total={tot_col}")

    for c in [lowg_col, deo_col, pcmo_col, tot_col]:
        if c:
            lcc[c] = pd.to_numeric(lcc[c], errors='coerce').fillna(0)

    # LOW GRADE volume is in litres; convert to ML
    lcc['lowg_ml'] = lcc[lowg_col] / 1_000_000 if lowg_col else 0
    lcc['deo_ml']  = lcc[deo_col]  / 1_000_000 if deo_col  else 0
    lcc['tot_ml']  = lcc[tot_col]              if tot_col  else 0  # already ML

    # Opportunity: shift LOW GRADE -> DEO gains PKR (348.1 - 116.0) = 232.1 per litre
    DEO_MGLN  = 348.1
    LOWG_MGLN = 116.0
    MGN_GAP   = DEO_MGLN - LOWG_MGLN  # PKR 232.1 / litre

    lcc['lowg_pct']        = (lcc['lowg_ml'] / lcc['tot_ml'].replace(0, float('nan')) * 100).round(1)
    lcc['mix_opp_PKR_M']   = (lcc['lowg_ml'] * MGN_GAP).round(1)   # PKR millions (ML * PKR/ltr = PKR M)

    top15 = lcc.nlargest(15, 'mix_opp_PKR_M')[[city_col, 'tot_ml', 'lowg_ml', 'lowg_pct', 'mix_opp_PKR_M']].copy()
    top15.columns = ['City', 'Total Lube (ML)', 'LOW GRADE (ML)', 'LOW GRADE %', 'Opp PKR M']
    top15 = top15.reset_index(drop=True)
    top15.index += 1
    print("\n  Top 15 cities — Lube Mix Opportunity (shift LOW GRADE -> DEO @ PKR 232/ltr)")
    print(top15.to_string())
    print(f"\n  TOTAL Lube Mix Opportunity (top 15): PKR {top15['Opp PKR M'].sum():.0f} M")

# ── B. FUEL DISCOUNT NORMALISATION OPPORTUNITY ───────────────────────────────
print("\n== B. FUEL DISCOUNT OPPORTUNITY vs South Benchmark ==")

SOUTH_PETROL_DISC = 0.28  # PKR/ltr
SOUTH_DIESEL_DISC = 0.40  # PKR/ltr

for seg, tkey, south_disc in [
    ('Petrol', 'petrol_by_city', SOUTH_PETROL_DISC),
    ('Diesel', 'diesel_by_city', SOUTH_DIESEL_DISC),
]:
    tbl = analysis_tables.get(tkey, pd.DataFrame())
    if tbl.empty:
        print(f"  {tkey} not found"); continue

    print(f"\n  {seg} city table columns: {list(tbl.columns[:8])}")

    # Find columns
    city_col  = 'CityNorm' if 'CityNorm' in tbl.columns else next(
        (c for c in tbl.columns if 'city' in c.lower()), None)
    vol_col   = next((c for c in tbl.columns if 'vol' in c.lower() and 'cy' in c.lower() and 'ml' in c.lower()), None)
    disc_col  = next((c for c in tbl.columns if 'disc' in c.lower() and 'ltr' in c.lower() and 'cy' in c.lower()), None)
    nmgn_col  = next((c for c in tbl.columns if 'nmgn' in c.lower() and 'ltr' in c.lower() and 'cy' in c.lower()), None)

    print(f"    city={city_col}  vol={vol_col}  disc={disc_col}  nmgn/ltr={nmgn_col}")
    if not all([city_col, vol_col, disc_col]):
        continue

    tbl = tbl.copy()
    tbl[vol_col]  = pd.to_numeric(tbl[vol_col],  errors='coerce').fillna(0)
    tbl[disc_col] = pd.to_numeric(tbl[disc_col], errors='coerce').fillna(0)

    tbl['excess_disc_ltr']  = (tbl[disc_col] - south_disc).clip(lower=0)
    tbl['disc_opp_PKR_M']   = (tbl['excess_disc_ltr'] * tbl[vol_col]).round(1)

    out = (tbl[tbl['excess_disc_ltr'] > 0]
           .nlargest(15, 'disc_opp_PKR_M')
           [[city_col, vol_col, disc_col, 'excess_disc_ltr', 'disc_opp_PKR_M']]
           .copy())
    out.columns = ['City', 'Vol CY (ML)', 'Disc/ltr (PKR)', 'Excess vs South', 'Opp PKR M']
    out = out.reset_index(drop=True)
    out.index += 1
    print(f"\n  Top 15 {seg} cities — Discount Normalisation Opportunity (South benchmark: PKR {south_disc}/ltr)")
    print(out.to_string())
    print(f"\n  TOTAL {seg} Discount Opp (top 15): PKR {out['Opp PKR M'].sum():.0f} M")

print("\nDone.")
