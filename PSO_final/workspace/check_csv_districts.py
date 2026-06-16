import pandas as pd
pumps = pd.read_csv('karachi_pumps.csv', dtype={'Code':str})
pumps['Code'] = pumps['Code'].str.strip().str.zfill(10)
pumps_kar = pumps[pumps['Division'].str.strip()=='Karachi'].copy()
names_to_check = ['Q. STAR','HUR','HAMSAFAR','KAUSAR','BISMILLAH','JOSH','A.REHMAN','AL-HAMD','NAWAZ','BHITTAI','ARABIA','RAANA','MEHRAN','AL-MADINA','AL MADINA','FAROOQ','CHASHMA','A.G','BHITTAI','BAHRIA','TARIQ','AWAMI','MODERN']
for n in names_to_check:
    rows = pumps_kar[pumps_kar['Name_of_Retail_Outlet'].str.contains(n, case=False, na=False)]
    for _, r in rows.iterrows():
        name = r['Name_of_Retail_Outlet'][:40]
        code = r['Code']
        dist = r['City_District_Area']
        print(f"{code}  {name:<40} {dist}")
print()
