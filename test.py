import pandas as pd
from DF_DBF import *
# столбцы для дэфолтного экспорта
from Export_columns import exp_format


def cross_columns_list(list_exist, list_target):
    """
    Возращаем список который входит в Exist
    cross_columns = cross_columns_list(total_columns, exp_format)

    list_exist - столбцы в наличии
    list_target - желаемый список столбцов

    cross_columns - только те из list_target что имеются в exist
    """

    z = 0
    cross_columns = []
    for val in list_target:
        for val2 in list_exist:
            z += 1
            if val == val2:
                cross_columns.append(val)
    print(z)
    return cross_columns


d = {'Шов': 7773, 'Вмятина': 14, 'Расслоеине': 119, 'Опора': 15, 'Отвод': 33, 'Трещина': 773, 'Маркер': 115,
     'Потеря металла': 10621}

c = {'Шов': 'WELD', 'Вмятина': 'ANOM', 'Потеря металла': 'ANOM', 'Расслоеине': 'ANOM', 'Опора': 'CON', 'Отвод': 'CON',
     'Трещина': 'ANOM', 'Гофр': 'ANOM', 'Маркер': 'MARKER'}

stat_df = pd.Series(data=d)
color_df = pd.Series(data=c)

tmp_df = stat_df.to_frame()
tmp_df = tmp_df.rename(columns={0: "index_column"})
tmp_df = tmp_df.reset_index()
tmp_df['tag'] = ''

for i, row in tmp_df.iterrows():
    tmp_df.at[i, 'tag'] = color_df[tmp_df.loc[i]['index']]

tmp_df.sort_values(['tag', 'index_column'], ascending=[True, False], inplace=True)

if __name__ == '__main__':
    df_dbf = df_DBF()

    a, b = df_dbf.parse_columns(columns_list=exp_format)

    print(a)
    print(b)
    # print(isml)
    # process_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
    # process_columns_df = process_columns_df.append({'COL_INDEX': 1, 'COL_NAME': 11}, ignore_index=True)
    # process_columns_df = process_columns_df.append({'COL_INDEX': 2, 'COL_NAME': 22}, ignore_index=True)
    # process_columns_df = process_columns_df.append({'COL_INDEX': 3, 'COL_NAME': 33}, ignore_index=True)
    #
    # process_columns_df.loc[-0.5] = [9, 99]  # adding a row
    # process_columns_df.index = process_columns_df.index + 1  # shifting index
    # process_columns_df = process_columns_df.sort_index()  # sorting by index
