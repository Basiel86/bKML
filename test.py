import pandas as pd

index_filename = ('DBF_INDEX.xlsx')

ind = 102

df_index_fea_code = pd.read_excel(index_filename, sheet_name='FEA_CODE')
id_list = df_index_fea_code['ID']

if not id_list[id_list.isin([ind])].empty:
    fea_ru = df_index_fea_code.loc[df_index_fea_code['ID'] == ind]
    print('found: ', fea_ru)
else:
    print("value not found")
