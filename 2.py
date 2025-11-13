import pandas as pd

# 1️⃣ Load Excel file
file_path = "cleaned_jizzax_turarjoy.xlsx"
df = pd.read_excel(file_path)

# 2️⃣ Define a helper function to insert a new column after a specific one
def insert_after(df, after_col, new_col, values):
    cols = df.columns.tolist()
    pos = cols.index(after_col) + 1
    df.insert(pos, new_col, values)
    return df

# 3️⃣ Create and insert each new column after its 4.x.2 column
df = insert_after(df, "4.1.2", "poydevor_cnt", df["4.1.1"] * df["4.1.2"])
df = insert_after(df, "4.2.2", "devor_cnt", df["4.2.1"] * df["4.2.2"])
df = insert_after(df, "4.3.2", "tomyopish_cnt", df["4.3.1"] * df["4.3.2"])
df = insert_after(df, "4.4.2", "kommunikatsiya_cnt", df["4.4.1"] * df["4.4.2"])

# 4️⃣ Save updated file
output_path = "cleaned_jizzax_turarjoy_add_cnt.xlsx"
df.to_excel(output_path, index=False)

print("✅ New columns added successfully and saved as:", output_path)