import pandas as pd
import langid
import xlsxwriter
from langdetect import detect
df_old = pd.read_excel("for_translation.xlsx")
#dfc = pd.read_csv("raw.csv")
#dfc.to_excel('raw.xlsx',encoding='utf-8',index = False)
df = pd.read_excel('raw.xlsx')
bengali_words = ['bengali']
english_words = ['english']
column_name = ['column_name']
lang = ['language']
row = ['respid']
for col in df.columns:
    if col == 'begin_group:respid' or col == "begin_group-respid":
        df.rename(columns = {col:'respid'}, inplace = True)
for column in df.columns:
    for word in df[column]:
        if type(word) == int or type(word) == float or type(word) == bool or 'uuid' in word or 'UTC' in word or '_' in word:
            continue
        lan = langid.classify(word)
        lanlast = langid.classify(word.split(" ")[-1])
        if lan[1] > 0:
            continue
        try:
            if lan[0] != 'en' or detect(word) != 'en' or (lan[0] == 'en' and lanlast[0]!= 'en') or (detect(word) == 'en' and detect(word.split(" ")[-1])!= 'en'):
                row.append(df[df[column] == word].respid.values)
                bengali_words.append(word)  
                column_name.append(column)
                lang.append(lan)
        except:
            row.append(df[df[column] == word].respid.values)
            bengali_words.append("ERROR!!! "+ word)  
            column_name.append(column)
            lang.append(lan)
english_words[0] = "english"
english_words[1:] = ["=IFERROR(INDEX(OLD!C$2:C$"+str(len(df_old)+1)+",MATCH(1,(NEW!B2=OLD!B$2:B$"+str(len(df_old)+1)+")*1,0),0),COUNT(C$1:C1)+1)" for nam in range(len(bengali_words))]
df = pd.DataFrame(zip(column_name,bengali_words,english_words,lang,row))
df_old.columns = ['column_name','bengali','english','language','respid']
main = pd.ExcelWriter("for_translation.xlsx", engine = 'xlsxwriter')
df.to_excel(main,sheet_name='NEW',encoding='utf-08',index=False,header = False)
df_old.to_excel(main,sheet_name='OLD',encoding='utf-08',index=False,header = True)
main.save()
print("FILTER HAS FINISHED RUNNING")
