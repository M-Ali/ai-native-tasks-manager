"""
Parse 25-08-2021-List-of-ROs.pdf to extract Karachi station addresses.
Format: Sr# | CITY | OUTLET NAME | ADDRESS
"""
import sys, re, json
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781498053596-dz5zjs.pdf'
r = pypdf.PdfReader(PDF)
pages = r.pages
print(f"Pages: {len(pages)}")

text = '\n'.join(p.extract_text() or '' for p in pages)
print(f"Chars: {len(text)}")

# Show first 3000 chars to understand format
print("\n=== FIRST 3000 CHARS ===")
print(text[:3000])

print("\n\n=== SEARCHING FOR KARACHI STATIONS ===")
# Find lines with Karachi
lines = text.split('\n')
kar_start = None
for i,l in enumerate(lines):
    if 'Karachi' in l or 'KARACHI' in l:
        kar_start = i
        break

if kar_start:
    print(f"Karachi section starts at line {kar_start}")
    print('\n'.join(lines[max(0,kar_start-2):kar_start+50]))
