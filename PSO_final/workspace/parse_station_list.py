"""
Parse PSO list_of_stations.pdf: Name Address City District Province Phone (all concatenated).
Extract Karachi stations with their GOS districts.
"""
import sys, re, json, difflib
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF = r'C:\Users\Muhamad.Ali\.clone\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781498706919-gl2n6q.pdf'

PDF2 = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781498706919-gl2n6q.pdf'

r = pypdf.PdfReader(PDF2)
full_text = ''.join(p.extract_text() or '' for p in r.pages)

# Phone pattern: 0[0-9]{3}-[0-9]{7}
PHONE = r'0\d{3}-\d{7}'

# Province names in Pakistan
PROVINCES = r'(?:PUNJAB|SINDH|KPK|BALOCHISTAN|BALUCHISTAN|AJK|GILGIT|ISLAMABAD)'

# Parse all records: each record ends with Province+Phone
# Pattern: (any text)(PUNJAB|SINDH|...)(phone)
record_pat = re.compile(
    rf'(.+?)({PROVINCES})\s*({PHONE})',
    re.DOTALL
)

# Find all records
# First, let's split on phone numbers to get individual records
# Phone is always at the end of each record
phones = list(re.finditer(PHONE, full_text))
print(f"Total records (phone numbers found): {len(phones)}")

# Extract each record by splitting between phones
records = []
for i, m in enumerate(phones):
    phone_end = m.end()
    # Start of this record: just after previous phone (or start of text)
    if i == 0:
        rec_start = 0
    else:
        rec_start = phones[i-1].end()
    segment = full_text[rec_start:phone_end]
    records.append(segment.strip())

print(f"Extracted {len(records)} records")

# Now parse each record to extract: Name, Address, City, District, Province, Phone
# Record format: {Name}{Address}{City}{District}{Province}{Phone}
# Key: Province is PUNJAB/SINDH/etc, Phone is 0xxx-xxxxxxx
# City and District need to be extracted from before Province

def parse_record(seg):
    # Find province
    pm = re.search(rf'(PUNJAB|SINDH|KPK|BALOCHISTAN|BALUCHISTAN|AJK|GILGIT|ISLAMABAD)\s*({PHONE})', seg)
    if not pm:
        return None
    province = pm.group(1)
    phone = pm.group(2)
    before = seg[:pm.start()].strip()
    # For Karachi stations (SINDH province), extract city and district
    if province == 'SINDH':
        # Look for KARACHI, KORANGI, MALIR, KEAMARI before SINDH
        # Possible districts: KARACHI EAST/WEST/SOUTH/CENTRAL, KORANGI, MALIR, KEAMARI, THATTA, SAJAWAL, HYDERABAD, etc.
        dist_pat = re.search(
            r'(KARACHI\s*(?:EAST|WEST|SOUTH|CENTRAL)?|KORANGI|MALIR|KEAMARI|THATTA|SAJAWAL|HYDERABAD|SUKKUR|LARKANA|DADU|NAWABSHAH|MIRPURKHAS|UMERKOT|TANDO ALLAH YAR|SANGHAR|KHAIRPUR|JACOBABAD|KASHMORE|SHIKARPUR|GHOTKI|KAMBAR)',
            before
        )
        if dist_pat:
            dist = dist_pat.group(1).strip()
            # City is the word just before district
            city_area = before[:dist_pat.start()].strip()
            # Name is hard to separate from address, but let's take whole before as name+addr
            return {'name_addr': city_area, 'district': dist, 'province': province, 'phone': phone}
    return None

karachi_records = []
for seg in records:
    r2 = parse_record(seg)
    if r2 and r2['district'] and 'KARACHI' in r2['district']:
        karachi_records.append(r2)

print(f"\nKarachi district records found: {len(karachi_records)}")

# Now fuzzy-match against our unmatched stations
sys.path.insert(0,'src')
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
unmatched = kar_stns[~kar_stns['Customer Number'].isin(code_to_dist)]

# Strip common words for fuzzy matching
def norm(n):
    s = str(n).upper()
    for w in ['SERVICE STATION','FILLING STATION','PETROLEUM SERVICE','PETROLEUM SERVICES',
              'PETROL SERVICE','PETROLEUM','SERVICE','STATION','FILLING','P/S','F/S',
              'T/S','S/S','N/V','DFA','DFB','CF','PSO','(PVTLTD)','):','LTD','&','CO','AND',
              '.','  ',' ']:
        s = s.replace(w, ' ')
    return ' '.join(s.split())

# Build list of PDF records for matching
pdf_names = [(norm(r['name_addr'][:50]), r['district'], r['name_addr'][:80]) for r in karachi_records]

print("\nSample Karachi records from PDF:")
for na, d, full in pdf_names[:20]:
    print(f"  [{d}] {full[:70]}")

# For each unmatched station, find best match in PDF
print("\n=== FUZZY MATCH RESULTS ===")
extra_map = {}
for _, row in unmatched.sort_values('ml_cy', ascending=False).head(40).iterrows():
    code = row['Customer Number']
    name = row['name']
    norm_name = norm(name)

    best_score = 0; best_dist = None; best_full = None
    for pn, pd_dist, pfull in pdf_names:
        score = difflib.SequenceMatcher(None, norm_name, pn).ratio()
        if score > best_score:
            best_score = score
            best_dist = pd_dist
            best_full = pfull

    mark = '***' if best_score >= 0.55 else '   '
    print(f"{mark} {code} [{row['ml_cy']:.1f}ML] {name:<40} → score={best_score:.2f} [{best_dist}] {(best_full or '')[:50]}")
    if best_score >= 0.55 and best_dist:
        extra_map[code] = best_dist

print(f"\nMapped {len(extra_map)} stations via fuzzy match")
