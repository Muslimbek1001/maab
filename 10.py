import pandas as pd
import numpy as np
import os

# 1️⃣ Filenames (relative to project folder)
excel_file = "average_gisht_costs_five_columns.xlsx"
cities_file = "regions.txt"
output_file = "average_gisht_costs_with_cities.xlsx"

# 2️⃣ Read Excel
df = pd.read_excel(excel_file)

# 3️⃣ Columns to calculate avg_cost
cols = ['bir_kv_pishgan_gisht_cost', 'bir_kv_xom_gisht_cost', 'bir_kv_boshqa_gisht_cost']

# 4️⃣ Convert text to numeric safely (text → 0)
def to_numeric_safe(val):
    try:
        return float(val)
    except:
        return 0

for col in cols:
    df[col + "_num"] = df[col].apply(to_numeric_safe)

# 5️⃣ Count valid numeric values
df["valid_count"] = df[[c + "_num" for c in cols]].gt(0).sum(axis=1)

# 6️⃣ Calculate avg_cost
df["avg_cost"] = (
    df[[c + "_num" for c in cols]].sum(axis=1) /
    df["valid_count"].replace(0, np.nan)
).round(2)

# 7️⃣ Drop helper columns
df.drop(columns=[c + "_num" for c in cols] + ["valid_count"], inplace=True)

# 8️⃣ Read cities.txt and create mapping
city_map = {}
if os.path.exists(cities_file):
    with open(cities_file, "r", encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split(maxsplit=1)
            if len(parts) == 2:
                code, name = parts
                city_map[int(code)] = name
else:
    raise FileNotFoundError(f"{cities_file} not found!")

# 9️⃣ Replace COATO codes with city names
df["COATO"] = df["COATO"].map(city_map).fillna(df["COATO"])

# 🔟 Save final Excel
df.to_excel(output_file, index=False)

print(f"✅ Done! Saved as {output_file}")
print(df.head())