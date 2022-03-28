import pandas as pd
import xlsxwriter
dictionary = pd.read_excel('dictionary.xlsx')
df_translated = pd.read_excel('for_translation.xlsx')
df_raw = pd.read_excel('raw.xlsx')
dic = pd.Series(df_translated['english'].values,index = df_translated['bengali']).to_dict()

df = pd.read_excel('raw.xlsx')


#df['column_name'] = df.column_name.replace(dic)
df = df.replace(dic)

main = pd.ExcelWriter('Translated_Covid_Survey_Week_00_results.xlsx',engine = 'xlsxwriter')

df.to_excel(main,encoding = 'utf-8',sheet_name='Translated',index =False)
df_raw.to_excel(main,encoding = 'utf-8',sheet_name='BeforeTranslation',index =False)

main.save()
dictionary = pd.concat([df_translated,dictionary],axis=0)
dictionary = dictionary.drop_duplicates(subset=['bengali'],keep='first',inplace=False,ignore_index=False)
dictionary = dictionary.drop(dictionary.columns[[0,3,4]],axis=1)
dictionary.to_excel('dictionary.xlsx',index=None)
print("REPLACE HAS FINISHED RUNNING")
