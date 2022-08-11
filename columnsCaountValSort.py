import os.path
import pandas as pd
import json

stat_file_name = 'columns_counts.json'


def load_columns_stat_dict():

    # открываем или создаем новый файл количества столбцов

    if os.path.exists(stat_file_name):
        try:
            columns_counts_dict = json.load(open(stat_file_name))
        except ValueError:
            print('# ERROR, JSON load failure! Created New Blank')
            columns_counts_dict = {}
    else:
        open(stat_file_name, 'a').close()
        columns_counts_dict = {}
    return columns_counts_dict


# экспортируем обратно в Json
def export_dict_to_json(columns_counts_dict):
    with open(stat_file_name, 'w') as file:
        file.write(json.dumps(columns_counts_dict, indent=3))  # use `json.loads` to do the reverse


def columns_count_val_sort(columns_list):

    # функция подсчета повторяющихся элементов в списке
    # возвращаем сортированный список всех используемых ранее столбцов

    # загружаем или открываем файл
    columns_counts_dict = load_columns_stat_dict()
    # делаем датафрейм из списка
    columns_df = pd.DataFrame(columns_list)
    # делаем статистику по стаобцам
    columns_df_stat = columns_df[0].value_counts(dropna=False, normalize=False)

    # бежим по списку столбцов
    for index, value in columns_df_stat.items():
        # если столбец есть в словаре, то инкремент
        if index in columns_counts_dict:
            columns_counts_dict[index] = columns_counts_dict[index] + 1
        else:
            # если нет то создаем новый ключ
            if index != "#BLANK":
                columns_counts_dict[index] = 1

    # сортируем словарь по количеству значени
    columns_counts_dict = dict(sorted(columns_counts_dict.items(), key=lambda item: item[1], reverse=True))

    # экспортируем обратно в файл
    export_dict_to_json(columns_counts_dict=columns_counts_dict)


