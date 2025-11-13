import pandas as pd

# Load your Excel file
df = pd.read_excel("all_xarajatlar.xlsx")

# List of required columns (keep the order you want)
required_columns = [
    'COATO','кв.м.',
    'poydevor_cnt','poydevor_xarajati',
    'devor_cnt','devor_xarajati',
    'tomyopish_cnt','tomyopish_xarajati',
    'kommunikatsiya_cnt','kommunikatsiya_xarajati',
    'Sement_xarajati','Tuproq_xarajati','Qumloq_xarajati','Sheben_xarajati',"Shag'al_xarajati",
    'Alebastr_xarajati','Tijorat beton_xarajati','Armatura_xarajati',
    'Yog\'och_xarajati','Pishgan g\'isht_xarajati','Xom g\'isht_xarajati','Shlakli g\'isht_xarajati','Shifer_xarajati',
    'Tunuka shifer_xarajati','Boshqa tom yopish materiallari_xarajati','Yog\'och oyna romlari_xarajati','Plastik oyna romlari_xarajati',
    'Metal oyna romlari_xarajati','Yog\'och eshiklar_xarajati','Plastik eshiklar_xarajati','Metak eshiklar_xarajati',
    'Temir darvoza_xarajati','Yog\'och darvoza_xarajati','Qora qog\'oz_xarajati','Shishali oyna_xarajati',
    'Turli diametrli plastik quvurlar_xarajati','Turli diametrli metal quvurlar_xarajati',
    'Turli o\'lchamdagi simlar_xarajati','Tabiiy tosh_xarajati','Gipskarton_xarajati','Lineolium_xarajati',
    'Keramik plitalar_xarajati','Bo\'yoq mahsulotlari_xarajati'
]

# Keep only the required columns (ignore if some columns are missing)
df_filtered = df[[col for col in required_columns if col in df.columns]]

# Save the filtered table
df_filtered.to_excel("selected_columns.xlsx", index=False)

print("✅ Filtered Excel file saved as 'selected_columns.xlsx'")