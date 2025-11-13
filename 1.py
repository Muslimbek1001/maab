import pandas as pd

columns = [
    'СОАТО', 'Код', 'Кадастр', 'Дата', 'кв.м.',
    '4.1.1', '4.1.2',
    '4.2.1', '4.2.2',
    '4.3.1', '4.3.2',
    '4.4.1', '4.4.2',
    '5', '6', '7.01', '7.02', '7.03', '7.04', '7.05', '7.06', '7.07', '7.08', '7.09', '7.10', '7.11', '7.12', '7.13',
    '7.14',
    '7.15', '7.16.1', '7.16.2', '7.16.3',
    '7.17.1', '7.17.2', '7.17.3',
    '7.18.1', '7.18.2', '7.19', '7.20', '7.21', '7.22', '7.23', '7.24', '7.25', '7.26', '7.27', '7.28', '7.29',
    '7.30'
]

# Excel faylni o'qish, header yo'q, ustun nomlari qo'yilgan
df = pd.read_excel("jizzax_turarjoy.xlsx", header=None, names=columns)

# Birinchi qatordagi ma'lumotni olib tashlash
df = df.drop(index=0).reset_index(drop=True)

print(df.columns)
# Natijani yangi Excel faylga saqlash
df.to_excel("cleaned_jizzax_turarjoy.xlsx", index=False)