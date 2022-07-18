import csv

import pandas as pd
import numpy as np
import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


struct_filename = resource_path('STRUCT.xlsx')

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

    cname = input("Enter columns: ")

    cname = cname.split('\t')

    cname = parse_columns(columns_list=cname, lang=lang)

    print(cname[1])


