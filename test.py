import pandas as pd

if __name__ == '__main__':

    index_filename = ('DBF_INDEX.xlsx')

    ind = 103

    df_index_fea_code = pd.read_excel(index_filename, sheet_name='FEA_CODE')
    id_list = df_index_fea_code['ID']

    for i, row in df_index_fea_code.iterrows():
        print(df_index_fea_code.loc[i]['ID'])

    if not id_list[id_list.isin([ind])].empty:
        fea_ru = df_index_fea_code.loc[df_index_fea_code['ID'] == ind]
        print('found: ', fea_ru)
        print(fea_ru.loc[1]['FEA_RU'], " / ", fea_ru.loc[1]['TYPE_RU'])

    else:
        print("value not found")
