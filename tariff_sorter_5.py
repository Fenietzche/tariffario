# -*- coding: utf-8 -*-
"""
Created on Sat Dec  4 16:26:51 2021

@author: Fenietzche
"""

import pandas
import datetime

def remove_row(d, lista):
    res = d.copy()
    for i in lista:
        res = res[res.index != i]
    res.index = list(range(len(res)))
    return res

def clear_outdated(dat):
    oggi = datetime.date.today()
    lista = []
    for i in dat.index:
        if isinstance(dat.at[i, 'End'], datetime.datetime):
            data_fine = dat.at[i, 'End'].date()
            if (data_fine - oggi).days < 0:
                lista.append(i)
    return remove_row(dat, lista)

def clear_blanks(dat):
    lista = list(range(1, 6))
    i = 6
    while i < len(dat.index):
        if pandas.isna(dat.iat[i, 2]):
            lista.extend(range(i+2, i+7))
            i += 6
        i += 1
    return remove_row(dat, lista)

def ordina(dat, df_ordinamento):
    d = dat.copy()
    for dest in d.index:
        if d.loc[str(dest), 'Codice'] in df_ordinamento.index:
            d[dest, 'Destinazione'] = df_ordinamento.loc[dest, 'Destinazione']
    d.sort_index(inplace=True)    
    return d


path_name = "rateshit.xls" #input("Inserisci il nome del file: ")
df = pandas.read_excel(path_name, sheet_name = 0)

df.columns = ['RateID', 'Start', 'End', 'Origin', 'Codice', 'Min', 'N', '+100kg', '+300kg', '+500kg', '+1000kg', 'Blank1', 'Blank2', 'Base']

df = clear_outdated(df)
df = clear_blanks(df)


df_sorted = []
i = 0
lung_df = len(df.index)
while i < lung_df:
    if pandas.isna(df.iat[i, 2]):
        j = i+1
        while not pandas.isna(df.iat[j, 2]):
            if j == lung_df-1:
                break
            j += 1
        
        df_appl = pandas.DataFrame(df.loc[i+1:j-1, 'Codice': 'Base'])
        list_dest = ['' for x in range(len(df_appl))]
        list_app = [df.iat[i,0] for x in range(len(df_appl))]
        df_iniz = pandas.DataFrame([list_dest, list_app]).transpose()
        df_iniz.columns = ['Destinazione', 'Applicability']
        df_appl = pandas.concat((df_iniz, df_appl), axis = 1)
        df_sorted.append(df_appl)
        i = j
    i += 1


writer = pandas.ExcelWriter('test3.xlsx', engine='openpyxl')
for i in range(len(df_sorted)):
    df_sorted[i].to_excel(writer, sheet_name=df_sorted[i].iat[0, 1], index=False)
writer.save()
writer.close()

        

