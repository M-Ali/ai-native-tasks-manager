import sys, pandas as pd
sys.path.insert(0,'src')
sys.stdout.reconfigure(encoding='utf-8')
from pso import ingest

df, _ = ingest.load('data/input/Working File Retail Fuels Data.xlsx')
retail = df[df['IsRetail'] & ~df['IsInternational']].copy()
kar = retail[retail['CityNorm']=='Karachi'].copy()
kar['Customer Number'] = kar['Customer Number'].astype(str).str.strip()

pumps = pd.read_csv('karachi_pumps.csv', dtype={'Code':str})
pumps['Code'] = pumps['Code'].str.strip().str.zfill(10)
pumps_kar = pumps[pumps['Division'].str.strip()=='Karachi'].copy()

# check for duplicate codes in CSV
dups = pumps_kar[pumps_kar['Code'].duplicated(keep=False)]
print(f'Duplicate codes in CSV: {len(dups)}')
print(dups[['Code','Name_of_Retail_Outlet','City_District_Area']])

print()
# Test the exact dict construction from the script
code_to_dist = pumps_kar.set_index("Code")["City_District_Area"].to_dict()
print(f'code_to_dist size: {len(code_to_dist)}')

check = ['0000113354','0000115364','0000115469','0000116152']
for c in check:
    print(f'  {c} in code_to_dist: {c in code_to_dist}  -> {code_to_dist.get(c,"MISSING")}')

# Check what matches
kar_stns = kar.groupby('Customer Number', as_index=False).agg(
    name=('Name 1','first'), vol_cy=('SalesLtr_CY','sum'))
matched = kar_stns[kar_stns['Customer Number'].isin(code_to_dist)]
unmatched = kar_stns[~kar_stns['Customer Number'].isin(code_to_dist)]
print(f'\nMatched: {len(matched)}  Unmatched: {len(unmatched)}')
print('\nUnmatched codes that ARE in code_to_dist:')
for _, r in unmatched.iterrows():
    if r['Customer Number'] in code_to_dist:
        print(f'  BUG: {r["Customer Number"]} IS in dict but shows as unmatched!')
