import pandas as pd
import numpy as np
import openpyxl as op
from datetime import timedelta
import pytz

# Место, куда сохранил файл из письма:
file = "C:/Users/Noizz/Desktop/learn/Excel отчёт/Сводка по объектам SM.xlsx"
df = pd.read_excel(file)

df.reset_index(drop= True, inplace= True)
df.rename(columns = {'НОМЕР ЗНО':'НОМЕР'}, inplace = True)
df.pop('Unnamed: 0')

df1 = pd.read_excel(file, sheet_name = 'ЗНР')
df1.rename(columns = {'НОМЕР ЗНР':'НОМЕР'}, inplace = True )
df1.pop('Unnamed: 0')

df2 = pd.read_excel(file, sheet_name = 'Инциденты')
df2.rename(columns = {'НОМЕР ИНЦИДЕНТА':'НОМЕР'}, inplace = True )
df2.pop('Unnamed: 0')

df3 = pd.read_excel(file, sheet_name = 'Мероприятия по рискам')
df3.rename(columns = {'НОМЕР МПР':'НОМЕР'}, inplace = True )
df3.pop('Unnamed: 0')

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)
pd.options.display.expand_frame_repr = False

df['СОТРУДНИК'] = df['СОТРУДНИК'].replace(np.nan, '----')
df1['СОТРУДНИК'] = df1['СОТРУДНИК'].replace(np.nan, '----')
df2['СОТРУДНИК'] = df2['СОТРУДНИК'].replace(np.nan, '----')
df3['СОТРУДНИК'] = df3['СОТРУДНИК'].replace(np.nan, '----')

df.sort_values(by = ["ОБЪЕКТ", "КОНТР СРОК"])

n = 48
date_start = pd.to_datetime('today').normalize()
date_end = date_start + timedelta(hours=n)

df_res = pd.concat([df, df1, df2, df3], ignore_index=True)
rslt_df = df_res[df_res["КОНТР СРОК"] <= date_end]

rslt_df = rslt_df[['НОМЕР','ОБЪЕКТ','КОНТР СРОК','ГРУППА','СОТРУДНИК']]
# Место, куда сохраняем готовый файл:
rslt_df.to_excel (r"C:/Users/Noizz/Desktop/learn/Excel отчёт/Otchet_SM.xlsx")
