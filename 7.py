import pandas as pd

# 1️⃣ Read Excel file
file_path = "selected_columns.xlsx"  # <-- change to your file
df = pd.read_excel(file_path)

# 2️⃣ List of common columns for sums (excluding specific brick columns)
common_cols = [
    "Sement_xarajati", "Tuproq_xarajati", "Qumloq_xarajati", "Sheben_xarajati",
    "Shag'al_xarajati", "Alebastr_xarajati", "Tijorat beton_xarajati", "Armatura_xarajati",
    "Yog'och_xarajati", "Shifer_xarajati", "Tunuka shifer_xarajati",
    "Boshqa tom yopish materiallari_xarajati", "Yog'och oyna romlari_xarajati",
    "Plastik oyna romlari_xarajati", "Metal oyna romlari_xarajati", "Yog'och eshiklar_xarajati",
    "Plastik eshiklar_xarajati", "Metak eshiklar_xarajati", "Temir darvoza_xarajati",
    "Yog'och darvoza_xarajati", "Qora qog'oz_xarajati", "Shishali oyna_xarajati",
    "Turli diametrli plastik quvurlar_xarajati", "Turli diametrli metal quvurlar_xarajati",
    "Turli o'lchamdagi simlar_xarajati", "Tabiiy tosh_xarajati", "Gipskarton_xarajati",
    "Lineolium_xarajati", "Keramik plitalar_xarajati", "Bo'yoq mahsulotlari_xarajati",
    "kommunikatsiya_xarajati", "tomyopish_xarajati", "devor_xarajati", "poydevor_xarajati"
]

# 3️⃣ Helper function to calculate bir_kv_cost safely
def calc_bir_kv_cost(row, brick_col):
    if pd.isna(row.get(brick_col, 0)) or row.get(brick_col, 0) == 0:
        return "Qaytadan tekshiring"
    if pd.isna(row.get("кв.м.", 0)) or row.get("кв.м.", 0) == 0:
        return "Qaytadan tekshiring"
    cols_to_sum = common_cols + [brick_col]
    return row[cols_to_sum].sum() / row["кв.м."]

# 4️⃣ Calculate each brick cost column
df["bir_kv_pishgan_gisht_cost"] = df.apply(lambda row: calc_bir_kv_cost(row, "Pishgan g'isht_xarajati"), axis=1)
df["bir_kv_xom_gisht_cost"] = df.apply(lambda row: calc_bir_kv_cost(row, "Xom g'isht_xarajati"), axis=1)
df["bir_kv_boshqa_gisht_cost"] = df.apply(lambda row: calc_bir_kv_cost(row, "Shlakli g'isht_xarajati"), axis=1)

# 5️⃣ Keep only required columns
final_df = df[["COATO", "bir_kv_pishgan_gisht_cost", "bir_kv_xom_gisht_cost", "bir_kv_boshqa_gisht_cost"]]

# 6️⃣ Save to new Excel
output_file = "gisht_turlari.xlsx"
final_df.to_excel(output_file, index=False)
print(f"Done! Saved only selected columns to '{output_file}'")