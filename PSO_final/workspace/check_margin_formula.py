import sys
sys.path.insert(0, 'src')
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from pso import ingest

df, _ = ingest.load('data/input/Working File Retail Fuels Data.xlsx')
retail = df[df['IsRetail'] & ~df['IsInternational']].copy()
lubes  = retail[retail['FuelSegment'] == 'Lubricants'].copy()

# Show a sample row with all financial columns
fin_cols = ['SalesGRS_CY','Disc_CY','MarignPrimary_CY',
            'Margin/Ltr_CY','NetMargin_CY','NetMargin/Ltr_CY','SalesLtr_CY']
sample = lubes[lubes['SalesLtr_CY'] > 0][fin_cols].head(10)
print("=== SAMPLE ROWS ===")
print(sample.to_string())
print()

# Verify relationships
row = lubes[lubes['SalesLtr_CY'] > 100].iloc[0]
print("=== SINGLE ROW BREAKDOWN ===")
print(f"  SalesGRS_CY      (Gross Revenue)    : {row['SalesGRS_CY']:,.2f}")
print(f"  Disc_CY          (Discount)          : {row['Disc_CY']:,.2f}")
print(f"  MarignPrimary_CY (Primary Margin)    : {row['MarignPrimary_CY']:,.2f}")
print(f"  NetMargin_CY     (Net Margin)        : {row['NetMargin_CY']:,.2f}")
print(f"  SalesLtr_CY      (Volume Litres)     : {row['SalesLtr_CY']:,.2f}")
print()
print(f"  Margin/Ltr_CY    (in data)           : {row['Margin/Ltr_CY']:,.4f}")
print(f"  NetMargin/Ltr_CY (in data)           : {row['NetMargin/Ltr_CY']:,.4f}")
print()
print(f"  MarignPrimary / SalesLtr             : {row['MarignPrimary_CY']/row['SalesLtr_CY']:,.4f}")
print(f"  NetMargin / SalesLtr                 : {row['NetMargin_CY']/row['SalesLtr_CY']:,.4f}")
print()
print(f"  SalesGRS - Disc                      : {row['SalesGRS_CY']-row['Disc_CY']:,.2f}")
print(f"  SalesGRS - Disc - MarignPrimary      : {row['SalesGRS_CY']-row['Disc_CY']-row['MarignPrimary_CY']:,.2f}")
print(f"  Does GRS - Disc = NetMargin?         : {abs((row['SalesGRS_CY']-row['Disc_CY']) - row['NetMargin_CY']) < 1}")

# National check
print()
print("=== NATIONAL TOTALS CHECK ===")
vcy = lubes['SalesLtr_CY'].sum()
mgn = lubes['NetMargin_CY'].sum()
prim = lubes['MarignPrimary_CY'].sum()
grs = lubes['SalesGRS_CY'].sum()
disc = lubes['Disc_CY'].sum()
pre_mgn_pl = lubes['NetMargin/Ltr_CY'].mean()
print(f"  Gross Revenue CY             : PKR {grs/1e9:.4f} Bn")
print(f"  Discount CY                  : PKR {disc/1e9:.4f} Bn")
print(f"  Primary Margin CY            : PKR {prim/1e9:.4f} Bn")
print(f"  Net Margin CY                : PKR {mgn/1e9:.4f} Bn")
print(f"  Volume CY                    : {vcy/1e6:.4f} ML")
print()
print(f"  NetMargin / Volume           : PKR {mgn/vcy:.4f} / Litre")
print(f"  Mean of NetMargin/Ltr_CY col : PKR {pre_mgn_pl:.4f} / Litre")
print(f"  Margin/Ltr_CY mean           : PKR {lubes['Margin/Ltr_CY'].mean():.4f} / Litre")
print()
print(f"  GRS - Disc = ?NetMargin      : PKR {(grs-disc)/1e9:.4f} Bn  (vs NetMargin {mgn/1e9:.4f} Bn)")
