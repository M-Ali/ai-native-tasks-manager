"""
Parse list_of_stations.pdf and build district map for unmatched Karachi stations.
PDF format: Name Address Tehsil City District Province Phone (all concatenated)
"""
import sys, re, json, difflib
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781496621579-biw6e5.pdf'
r = pypdf.PdfReader(PDF)
text = ''.join(p.extract_text() or '' for p in r.pages)

# In this PDF, each entry ends with: ...{City}{District}SINDH{Phone}
# For Karachi stations, pattern: ...KARACHI{DISTRICT}SINDH{phone}
# where DISTRICT is one of: EAST, WEST, SOUTH, CENTRAL, KORANGI, MALIR, KEAMARI, (blank)
# But city and district are concatenated, e.g. "KARACHIKARACHISOUTH" or "KARACHIMALIR"

# Let's look at text around SINDH to understand the city+district pattern
sindh_positions = [m.start() for m in re.finditer(r'SINDH', text)]
print(f'SINDH occurrences: {len(sindh_positions)}')

# Show 20 examples to understand format
print('\nSample SINDH contexts:')
for pos in sindh_positions[:25]:
    context = text[max(0,pos-120):pos+25]
    print(f'  ...{context[-120:]}...')

print('\n\n=== Looking for KARACHI+EAST or WEST etc. ===')
kd_pat = re.compile(r'(KARACHI\s*(?:EAST|WEST|SOUTH|CENTRAL|KORANGI|MALIR|KEAMARI)?)\s*SINDH')
kd_matches = list(kd_pat.finditer(text))
print(f'Karachi+district+SINDH patterns: {len(kd_matches)}')
for m in kd_matches[:20]:
    ctx = text[max(0,m.start()-200):m.start()]
    print(f'  DIST=[{m.group(1).strip()}]  ...{ctx[-150:]}')
    print()
