"""
Search fuel card PDF for specific target names.
"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
import pypdf

PDF = r'C:\Users\Muhamad.Ali\.claude\projects\D--Personal-PSO-final\793854cf-2ae9-4b02-bbe2-4ee98082c598\tool-results\webfetch-1781499905640-u1c0tv.pdf'
r = pypdf.PdfReader(PDF)
text = '\n'.join(p.extract_text() or '' for p in r.pages)

TARGETS2 = [
    'ENERGY','BABAR','NASEEM','AHMAD','DARBAR','SHOZAB','RIND',
    'JOHAR','KISWA','NAWAZ 1','RAJA 3','ADVANCE',
    'NZ NORTH','NZ SOUTH','GULSHAN AND','ALI HUSSAIN',
    'IBRAHIM P','ASAD F','MAHIY','HAFIZ',
    'PSO S/S 76','PSO S/S 83','PSO S/S 85','PSO S/S 86','PSO S/S 87',
    'PSO S/S 91','PSO S/S 93','A&Y','A & Y',
    'FAISAL PS', 'PSO COCO (SD-1)', 'SD-1',
]

for t in TARGETS2:
    idx = text.upper().find(t.upper())
    if idx >= 0:
        ctx = text[max(0,idx-20):idx+200].replace('\n', ' ')
        print(f"FOUND [{t}]: {ctx.strip()[:160]}")
    else:
        print(f"NOT FOUND: {t}")
