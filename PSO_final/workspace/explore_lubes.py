import sys; sys.path.insert(0,'src')
sys.stdout.reconfigure(encoding='utf-8')
from pso import ingest
import pandas as pd

df, _ = ingest.load('data/input/Working File Retail Fuels Data.xlsx')
retail = df[df['IsRetail'] & ~df['IsInternational']].copy()
lubes = retail[retail['FuelSegment']=='Lubricants'].copy()

print('LUBE CATEGORIES:')
print(lubes['LubeCategory'].value_counts().to_string())
print()
print('REGIONS:')
print(lubes['Sales office Region'].value_counts().to_string())
print()
print('TOP 20 CITIES by CY volume (KL):')
city_vol = lubes.groupby('CityNorm').agg(
    vol_cy=('SalesLtr_CY','sum'), vol_ly=('SalesLtr_LY','sum')
).assign(chg=lambda d:(d.vol_cy-d.vol_ly)/d.vol_ly.abs()*100
).sort_values('vol_cy', ascending=False).head(20)
for c,(v,ly,chg) in city_vol.iterrows():
    print(f'  {c:<25} CY:{v/1000:>7.1f}KL  LY:{ly/1000:>7.1f}KL  Chg:{chg:>+6.1f}%')
print()
print('NATIONAL TOTALS:')
vc = lubes['SalesLtr_CY'].sum()
vl = lubes['SalesLtr_LY'].sum()
rc = lubes['SalesGRS_CY'].sum()
rl = lubes['SalesGRS_LY'].sum()
mc = lubes['NetMargin_CY'].sum()
ml = lubes['NetMargin_LY'].sum()
print(f'  Vol CY:    {vc/1e6:.3f} ML  |  LY: {vl/1e6:.3f} ML  |  Chg: {(vc-vl)/vl*100:+.1f}%')
print(f'  Rev CY:    PKR {rc/1e9:.3f} Bn  |  LY: {rl/1e9:.3f} Bn  |  Chg: {(rc-rl)/rl*100:+.1f}%')
print(f'  Margin CY: PKR {mc/1e9:.3f} Bn  |  LY: {ml/1e9:.3f} Bn  |  Chg: {(mc-ml)/ml*100:+.1f}%')
print(f'  Margin/Ltr CY: PKR {mc/vc:.2f}  |  LY: PKR {ml/vl:.2f}')
print(f'  Stations:  {lubes["Customer Number"].nunique()}')
print(f'  Cities:    {lubes["CityNorm"].nunique()}')
print()
print('BY CATEGORY:')
cat = lubes.groupby('LubeCategory').agg(
    vol_cy=('SalesLtr_CY','sum'), vol_ly=('SalesLtr_LY','sum'),
    rev_cy=('SalesGRS_CY','sum'), margin_cy=('NetMargin_CY','sum')
).assign(chg=lambda d:(d.vol_cy-d.vol_ly)/d.vol_ly.abs()*100,
         vol_sh=lambda d:d.vol_cy/d.vol_cy.sum()*100
).sort_values('vol_cy', ascending=False)
for cat_name, r in cat.iterrows():
    print(f'  {cat_name:<20} Vol:{r.vol_cy/1000:>7.1f}KL  {r.vol_sh:>5.1f}%  Chg:{r.chg:>+6.1f}%  Mgn/L:{r.margin_cy/r.vol_cy:.0f}')
print()
print('BY REGION:')
reg = lubes.groupby('Sales office Region').agg(
    vol_cy=('SalesLtr_CY','sum'), vol_ly=('SalesLtr_LY','sum'),
    margin_cy=('NetMargin_CY','sum'), stns=('Customer Number','nunique')
).assign(chg=lambda d:(d.vol_cy-d.vol_ly)/d.vol_ly.abs()*100,
         vol_sh=lambda d:d.vol_cy/d.vol_cy.sum()*100
).sort_values('vol_cy', ascending=False)
for rn, r in reg.iterrows():
    print(f'  {rn:<12} Vol:{r.vol_cy/1000:>7.1f}KL  {r.vol_sh:>5.1f}%  Chg:{r.chg:>+6.1f}%  Stns:{int(r.stns)}  Mgn/L:{r.margin_cy/r.vol_cy:.0f}')
print()
print('TOP 20 STATIONS by CY volume:')
stn = lubes.groupby(['Customer Number','Name 1','CityNorm','Sales office Region']).agg(
    vol_cy=('SalesLtr_CY','sum'), vol_ly=('SalesLtr_LY','sum'),
    margin_cy=('NetMargin_CY','sum')
).reset_index().assign(chg=lambda d:(d.vol_cy-d.vol_ly)/d.vol_ly.abs()*100
).sort_values('vol_cy', ascending=False).head(20)
for _, r in stn.iterrows():
    print(f'  {r["Name 1"][:30]:<30}  {r.CityNorm[:15]:<15}  {r.vol_cy/1000:>6.1f}KL  Chg:{r.chg:>+6.1f}%  Mgn/L:{r.margin_cy/r.vol_cy:.0f}')
