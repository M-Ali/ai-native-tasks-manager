import sys
sys.path.insert(0,'src')
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from pso import ingest

df, _ = ingest.load('data/input/Working File Retail Fuels Data.xlsx')
retail = df[df['IsRetail'] & ~df['IsInternational']].copy()
kar = retail[retail['CityNorm']=='Karachi'].copy()
kar['Customer Number'] = kar['Customer Number'].astype(str).str.strip()

pumps = pd.read_csv('karachi_pumps.csv', dtype={'Code':str})
pumps['Code'] = pumps['Code'].str.strip().str.zfill(10)
pumps_kar = pumps[pumps['Division'].str.strip()=='Karachi'].copy()
code_to_dist = pumps_kar.set_index('Code')['City_District_Area'].to_dict()

kar_stns = kar.groupby('Customer Number', as_index=False).agg(
    name=('Name 1','first'), vol_cy=('SalesLtr_CY','sum'))
kar_stns['ml_cy'] = kar_stns['vol_cy']/1e6
unmatched = kar_stns[~kar_stns['Customer Number'].isin(code_to_dist)].sort_values('ml_cy', ascending=False)

print(f'Total unmatched: {len(unmatched)}')
print(f'Total unmatched vol: {unmatched["ml_cy"].sum():.1f}ML')
print()
print('ALL unmatched stations:')
for _, r in unmatched.iterrows():
    print(f'{r["Customer Number"]}  {r["name"]:<50} {r["ml_cy"]:.2f}ML')
