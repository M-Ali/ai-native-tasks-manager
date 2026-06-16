"""
Parse fuel_card PDF - get ALL Karachi entries.
"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781499905640-u1c0tv.pdf'
r = pypdf.PdfReader(PDF)
text = '\n'.join(p.extract_text() or '' for p in r.pages)

# Print all Karachi entries (lines mentioning KARACHI)
lines = text.split('\n')
print(f"Total lines: {len(lines)}")
print("\n=== ALL LINES WITH KARACHI ===")
for i, l in enumerate(lines):
    if 'KARACHI' in l.upper() and len(l.strip()) > 10:
        # Also show context
        print(f"Line {i}: {l.strip()}")
