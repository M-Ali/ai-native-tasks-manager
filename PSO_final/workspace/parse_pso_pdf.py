"""
Parse PSO premier-euro-5.pdf and participating-outlets.pdf
Build a comprehensive Code -> District map and save for use in the district slide.
"""
import sys, re, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF1 = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781495503036-2d1kbb.pdf'
# participating outlets pdf
PDF2 = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781495501613-g0t7rg.pdf'
# ro list (2021)
PDF3 = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781493542659-1rpzee.pdf'

def extract(path):
    r = pypdf.PdfReader(path)
    return '\n'.join(p.extract_text() or '' for p in r.pages)

text1 = extract(PDF1)
text2 = extract(PDF2)
text3 = extract(PDF3)

print(f"PDF1 chars: {len(text1)}")
print(f"PDF2 chars: {len(text2)}")
print(f"PDF3 chars: {len(text3)}")

# ── Parse premier-euro-5.pdf ─────────────────────────────────────────────────
# Format: Sr# Code Division Name SiteCat District Address (space-separated, tricky)
# Pattern: line starts with integer, then 6-digit code, then city, then rest
# Example: "1 101535 Karachi AHMED SERVICE STATION N/V CF KARACHI EAST DADABHOY..."

rows1 = []
# Each record starts with a number followed by a 6-digit code
lines1 = text1.replace('\n',' ')
# Split on pattern: digit(s) followed by 6-digit number followed by known city
pattern = r'(\d+)\s+(\d{6})\s+(Karachi|Lahore|Multan|Islamabad|Hyderabad|Jhelum|Gujranwala|Faisalabad|Rawalpindi|Quetta|Peshawar)\s+'
matches = list(re.finditer(pattern, lines1))
for i, m in enumerate(matches):
    end = matches[i+1].start() if i+1 < len(matches) else len(lines1)
    segment = lines1[m.end():end].strip()
    # Segment: "AHMED SERVICE STATION N/V CF KARACHI EAST DADABHOY NOROJI ROAD..."
    # Known site cats: N/V CF, N/V DFA, N/V DFB, S/S CF, S/S DFA, S/S DFB, T/S DFA, T/S DFB, F/S DFA, COCO SITES
    site_pat = r'\s+(N/V CF|N/V DFA|N/V DFB|S/S CF|S/S DFA|S/S DFB|T/S DFA|T/S DFB|F/S DFA|F/S DFB|COCO SITES)\s+'
    sm = re.search(site_pat, segment)
    if sm:
        name = segment[:sm.start()].strip()
        after_site = segment[sm.end():].strip()
        # after_site: "KARACHI EAST DADABHOY NOROJI ROAD..."
        # District is KARACHI EAST, KARACHI WEST, KARACHI SOUTH, MALIR, THATTA, LASBELA, etc.
        dist_pat = r'^(KARACHI EAST|KARACHI WEST|KARACHI SOUTH|KARACHI CENTRAL|MALIR|THATTA|LASBELA|SAJAWAL|NAWABSHAH|GWADER|KARACHI)\s+'
        dm = re.match(dist_pat, after_site)
        dist = dm.group(1) if dm else 'KARACHI'
        addr = after_site[dm.end():].strip() if dm else after_site
    else:
        name = segment[:40].strip()
        dist = 'KARACHI'
        addr = segment[40:].strip()
    rows1.append({
        'sr': m.group(1), 'code': m.group(2).zfill(10),
        'division': m.group(3), 'name': name,
        'district': dist, 'address': addr[:80]
    })

df1 = pd.DataFrame(rows1)
kar1 = df1[df1['division']=='Karachi']
print(f"\nPDF1 (premier-euro5): {len(df1)} total, {len(kar1)} Karachi")
print(kar1[['code','name','district']].head(10).to_string())

# Save the Karachi ones
kar1.to_csv('workspace/pso_pdf_stations.csv', index=False)
print(f"\nSaved {len(kar1)} Karachi stations to workspace/pso_pdf_stations.csv")

# ── Now match against our 92 unmatched stations ──────────────────────────────
sys.path.insert(0,'src')
from pso import ingest
df_main, _ = ingest.load('data/input/Working File Retail Fuels Data.xlsx')
retail = df_main[df_main['IsRetail'] & ~df_main['IsInternational']].copy()
kar_main = retail[retail['CityNorm']=='Karachi'].copy()
kar_main['Customer Number'] = kar_main['Customer Number'].astype(str).str.strip()

pumps = pd.read_csv('karachi_pumps.csv', dtype={'Code':str})
pumps['Code'] = pumps['Code'].str.strip().str.zfill(10)
pumps_kar = pumps[pumps['Division'].str.strip()=='Karachi'].copy()

# Original code map from CSV
csv_code_map = pumps_kar.set_index('Code')['City_District_Area'].to_dict()

# New code map from PDF1
pdf_code_map = {r['code']: r['district'] for _, r in kar1.iterrows()}

# Combined
combined_map = {**pdf_code_map, **csv_code_map}  # CSV takes priority
print(f"\nCSV codes: {len(csv_code_map)}")
print(f"PDF codes: {len(pdf_code_map)}")
print(f"Combined unique: {len(combined_map)}")

# Station-level totals
kar_stns = (kar_main.groupby('Customer Number', as_index=False)
            .agg(name=('Name 1','first'),
                 vol_cy=('SalesLtr_CY','sum'),
                 vol_ly=('SalesLtr_SPLY','sum'))
            .assign(ml_cy=lambda d: d['vol_cy']/1e6,
                    ml_ly=lambda d: d['vol_ly']/1e6))

matched   = kar_stns[kar_stns['Customer Number'].isin(combined_map)]
unmatched = kar_stns[~kar_stns['Customer Number'].isin(combined_map)]
print(f"\nWith combined map -> Matched: {len(matched)}  Unmatched: {len(unmatched)}")
print(f"Unmatched vol: {unmatched['ml_cy'].sum():.1f}ML")

print("\nRemaining unmatched (not in CSV or PDF) — top 20 by vol:")
for _, r in unmatched.sort_values('ml_cy', ascending=False).head(20).iterrows():
    print(f"  {r['Customer Number']}  {r['name']:<45} {r['ml_cy']:.1f}ML")
