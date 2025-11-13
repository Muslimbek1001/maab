import pandas as pd

# 1️⃣ Load Excel file
df = pd.read_excel('combined_data.xlsx')

# 2️⃣ Define materials and corresponding codes (must match column names in Excel)
materials = [
    "Sement", "Tuproq", "Qumloq", "Sheben", "Shag'al", "Alebastr", "Tijorat beton", "Armatura", "Yog'och",
    "Pishgan g'isht", "Xom g'isht", "Shlakli g'isht", "Shifer", "Tunuka shifer", "Boshqa tom yopish materiallari",
    "Yog'och oyna romlari", "Plastik oyna romlari", "Metal oyna romlari", "Yog'och eshiklar", "Plastik eshiklar",
    "Metak eshiklar", "Temir darvoza", "Yog'och darvoza", "Qora qog'oz", "Shishali oyna",
    "Turli diametrli plastik quvurlar", "Turli diametrli metal quvurlar", "Turli o'lchamdagi simlar",
    "Tabiiy tosh", "Gipskarton", "Lineolium", "Keramik plitalar", "Bo'yoq mahsulotlari"
]

codes = [
    "7.01", "7.02", "7.03", "7.04", "7.05", "7.06", "7.07", "7.08", "7.09", "7.10", "7.11", "7.12",
    "7.13", "7.14", "7.15", "7.16.1", "7.16.2", "7.16.3", "7.17.1", "7.17.2", "7.17.3", "7.18.1",
    "7.18.2", "7.19", "7.20", "7.21", "7.22", "7.23", "7.24", "7.25", "7.26", "7.27", "7.28"
]

# 3️⃣ Calculate xarajat columns automatically
for mat, code in zip(materials, codes):
    new_col = f"{mat}_xarajati"
    if mat in df.columns and code in df.columns:   # Check both columns exist
        df[new_col] = df[mat] * df[code]
    else:
        print(f"⚠️ Missing columns: {mat} or {code}")

# 4️⃣ Save result to a new Excel file
df.to_excel('part_xarajatlar.xlsx', index=False)

print("✅ New *_xarajati columns added successfully!")
