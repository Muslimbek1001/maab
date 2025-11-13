import pandas as pd

# 1️⃣ Read Excel
file_path = "gisht_turlari.xlsx"
df = pd.read_excel(file_path)

# 2️⃣ Define cost columns
cost_cols = ['bir_kv_pishgan_gisht_cost', 'bir_kv_xom_gisht_cost', 'bir_kv_boshqa_gisht_cost']

# 3️⃣ Convert cities to 0 for calculation
df_numeric = df.copy()
for col in cost_cols:
    df_numeric[col] = pd.to_numeric(df_numeric[col], errors='coerce').fillna(0)

# 4️⃣ Group by COATO and calculate average
grouped = df_numeric.groupby("COATO")[cost_cols].mean().reset_index()

# 5️⃣ Replace 0 in averages with 'Qaytadan tekshiring'
for col in cost_cols:
    grouped[col] = grouped[col].apply(lambda x: "Qaytadan tekshiring" if x == 0 else round(x, 2))

# 6️⃣ Save result
output_file = "average_gisht_costs.xlsx"
grouped.to_excel(output_file, index=False)

print(f"Done! Grouped averages with 'Qaytadan tekshiring' saved to '{output_file}'")