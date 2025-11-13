import pandas as pd

# Load Excel file
df = pd.read_excel("part_xarajatlar.xlsx")

# Existing column names
columns = ['Bir kunlik ish haqi', 'poydevor_cnt', 'devor_cnt', 'tomyopish_cnt', 'kommunikatsiya_cnt']

# 1️⃣ Add new columns after the corresponding existing columns
# poydevor_xarajati = Bir kunlik ish haqi * poydevor_cnt
df.insert(df.columns.get_loc('poydevor_cnt') + 1,
          'poydevor_xarajati',
          df['Bir kunlik ish haqi'] * df['poydevor_cnt'])

# devor_xarajati = Bir kunlik ish haqi * devor_cnt
df.insert(df.columns.get_loc('devor_cnt') + 1,
          'devor_xarajati',
          df['Bir kunlik ish haqi'] * df['devor_cnt'])

# tomyopish_xarajati = Bir kunlik ish haqi * tomyopish_cnt
df.insert(df.columns.get_loc('tomyopish_cnt') + 1,
          'tomyopish_xarajati',
          df['Bir kunlik ish haqi'] * df['tomyopish_cnt'])

# kommunikatsiya_xarajati = Bir kunlik ish haqi * kommunikatsiya_cnt
df.insert(df.columns.get_loc('kommunikatsiya_cnt') + 1,
          'kommunikatsiya_xarajati',
          df['Bir kunlik ish haqi'] * df['kommunikatsiya_cnt'])

# Save the updated Excel
df.to_excel("all_xarajatlar.xlsx", index=False)

print("✅ New _xarajati columns added in specified positions!")