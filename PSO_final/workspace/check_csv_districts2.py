import sys, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
pumps = pd.read_csv('karachi_pumps.csv', dtype={'Code':str})
pumps['Code'] = pumps['Code'].str.strip().str.zfill(10)
pumps_kar = pumps[pumps['Division'].str.strip()=='Karachi'].copy()
names = ['SUPER SERVICE','GALAXY','HUR SERVICE','STANDARD SERVICE','BURRAQUE','KAARSAZ','SHEERIN','DARBAR','NASEEM','AHMAD','MEGRAN','PSO SERVICE STATION','PSO S/S','MEGHAN']
for n in names:
    rows = pumps_kar[pumps_kar['Name_of_Retail_Outlet'].str.contains(n, case=False, na=False)]
    for _, r in rows.iterrows():
        print(f"{r['Code']}  {r['Name_of_Retail_Outlet']:<42}  {r['City_District_Area']}")
print()
print('ALL PSO COCO stations in CSV:')
coco = pumps_kar[pumps_kar['Name_of_Retail_Outlet'].str.contains('PSO S/S|PSO SERVICE STATION|PSO COCO', case=False, na=False)]
for _, r in coco.iterrows():
    print(f"{r['Code']}  {r['Name_of_Retail_Outlet']:<42}  {r['City_District_Area']}")
