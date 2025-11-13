import pandas as pd
import numpy as np

# 1️⃣ Read the Excel or CSV file
df = pd.read_excel("average_gisht_costs.xlsx")  # adjust path if needed

# 2️⃣ Column names
cols = ['bir_kv_pishgan_gisht_cost', 'bir_kv_xom_gisht_cost', 'bir_kv_boshqa_gisht_cost']

# 3️⃣ Function to check numeric and handle conversion
def to_numeric_safe(val):
    try:
        return float(val)
    except:
        return 0  # cities or invalid → 0

# 4️⃣ Apply conversion
for col in cols:
    df[col + "_num"] = df[col].apply(to_numeric_safe)

# 5️⃣ Calculate count of numeric values (non-zero numeric)
df["valid_count"] = df[[c + "_num" for c in cols]].gt(0).sum(axis=1)

# 6️⃣ Calculate avg_cost (sum / count)
df["avg_cost"] = (
    df[[c + "_num" for c in cols]].sum(axis=1) /
    df["valid_count"].replace(0, np.nan)
).round(2)

# 7️⃣ Optional: drop helper columns if not needed
df.drop(columns=[c + "_num" for c in cols] + ["valid_count"], inplace=True)

# 8️⃣ Save result
df.to_excel("average_gisht_costs_five_columns.xlsx", index=False)

print(df.head())