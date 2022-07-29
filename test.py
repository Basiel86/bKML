import csv

import pandas as pd
import numpy as np
import sys
import os

from difflib import SequenceMatcher
import jellyfish


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


struct_filename = resource_path(r'IDs\STRUCT.xlsx')

df_struct_col_var = pd.read_excel(struct_filename, sheet_name='COL_VAR')
df_struct_col_id_formats = pd.read_excel(struct_filename, sheet_name='COL_ID_STRUCT')


def parse_columns(columns_list, lang='RU'):
    col_id_list = []
    column_names = []

    if len(columns_list) > 0:
        for column_name in columns_list:
            cname_struct = '#BLANK'
            col_name = '#BLANK'
            col_id = '#BLANK'

            for i, row in df_struct_col_var.iterrows():
                if column_name == df_struct_col_var.loc[i]['COL_VAR_NAME']:
                    cname_struct = df_struct_col_var.loc[i]['COL_ID']
                elif cname_struct == '':
                    cname_struct = column_name

            for i, row in df_struct_col_id_formats.iterrows():
                if cname_struct == df_struct_col_id_formats.loc[i]['COL_ID']:
                    col_id = df_struct_col_id_formats.loc[i]['COL_ID']
                    col_name = df_struct_col_id_formats.loc[i][f'COL_NAME_{lang}']
                elif col_id == '':
                    col_id = cname_struct
                    col_name = cname_struct

            col_id_list.append(col_id)
            column_names.append(col_name)

    return col_id_list, column_names


if __name__ == '__main__':
    cname = ['Depth, mm', 'WTF1', "Orientation o'clock"]
    lang = 'RU'

    # cname = input("Enter columns: ")

    cname = '№ особ.	Дист., м	Отн. дист., м.	Длина секции, м	№ секции	Особенность	Описание	Комментарий	Т. ст., мм	Угол, °	Тип	Класс	ERF ASME B31G	P безоп. ASME B31G (МПа)	Длина, мм	Ширина, мм	Глуб., % от WT/Dн	Ост. т. ст., мм	Кластер	Срок ремонта особенности, лет	Широта, °	Долгота, °	Высота, м	Дист., м	Отн. дист., м.'

    cname = cname.split('\t')

    cname = parse_columns(columns_list=cname, lang=lang)

    a = 'Отн. дист м.'

    b = 'дист м.'
    c = 'дист. отн. м'

    print(jellyfish.levenshtein_distance(a, b))
    print(jellyfish.levenshtein_distance(a, c))

    print(jellyfish.hamming_distance(a, b))
    print(jellyfish.hamming_distance(a, c))

    print(similar(a, b))
    print(similar(a, c))

    # print(cname[1])
