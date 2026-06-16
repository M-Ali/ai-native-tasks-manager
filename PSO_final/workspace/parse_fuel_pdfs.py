"""
Parse fuel card PDFs to find addresses for unmatched Karachi stations.
"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDFs = {
    'fuel_card': r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781499905640-u1c0tv.pdf',
    'participating': r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781499903827-g9qyv9.pdf',
}

# Target stations to find
TARGETS = [
    'ENERGY PETROLEUM', 'BABAR FILLING', 'A&Y SERVICE', 'DARBAR FILLING',
    'NASEEM FILLING', 'AHMAD BROTHERS', 'IBRAHIM P/S', 'RIND', 'ASAD',
    'MAHIY SERVICE', 'SHOZAB', 'ADVANCE PETROLEUM', 'HAFIZ',
    'NAWAZ 1', 'RAJA 3', 'KISWA', 'GLOBAL PETROLEUM',
    'FAISAL SERVICE STATION', 'NZ NORTH', 'NZ SOUTH',
    'GULSHAN AND COMPANY', 'ALI HUSSAIN',
]

for name, path in PDFs.items():
    r = pypdf.PdfReader(path)
    text = '\n'.join(p.extract_text() or '' for p in r.pages)
    print(f"\n{'='*60}")
    print(f"PDF: {name} | {len(text)} chars")
    print(f"First 1000 chars:")
    print(text[:1000])
    print(f"\nSearching for target stations...")
    for t in TARGETS:
        idx = text.upper().find(t.upper())
        if idx >= 0:
            ctx = text[max(0,idx-20):idx+150]
            print(f"  FOUND [{t}]: {ctx.strip()[:130]}")
    print()
