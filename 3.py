import pandas as pd
df1=pd.read_excel('Jizzax_narx.xlsx')
df2=pd.read_excel('cleaned_jizzax_turarjoy_add_cnt.xlsx')
merged_df=pd.merge(df1,df2,left_on='COATO',right_on='СОАТО',how='left')
merged_df.to_excel('combined_data.xlsx',index='False')
print('Merged file saved')
