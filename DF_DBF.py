from tkinter import *
import os.path
import pandas as pd
import numpy as np
import sys
from simpledbf import Dbf5
from parse_inch import parse_inch_prj
from parse_inch import parse_inch_path
from dimm import dimm
import errno
from other_functions import dbf_feature_type_combine
from other_functions import dbf_description_combine

# столбцы для дэфолтного экспорта
from Export_columns import exp_format

index_filename_remote_path = r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\DBF_INDEX.xlsx'
index_filename_local_path = r'IDs\DBF_INDEX.xlsx'
struct_filename_remote_path = r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\STRUCT.xlsx'
struct_filename_local_path = r'IDs\STRUCT.xlsx'


def same_as_upper(col: pd.Series) -> pd.Series:
    '''
    Recursively fill NaN rows with the previous value
    '''
    if any(pd.Series(col).isna()):
        col = pd.Series(np.where(col.isna(), col.shift(1), col))
        return same_as_upper(col)
    else:
        return col


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Берем версию с компа или удаленно
def remote_or_local(remote_path, local_path):
    filename = os.path.basename(remote_path)[:-5]
    if os.path.exists(remote_path):
        print(f"# Info: {filename} file info: Remote")
        return remote_path
    else:
        print(f"# Info: {filename} file info: Local")
        return local_path


class df_DBF:
    def __init__(self):

        self.dbf_path = ''
        self.lng = 'RU'

        self.index_filename = remote_or_local(remote_path=index_filename_remote_path,
                                              local_path=index_filename_local_path)
        self.struct_filename = remote_or_local(remote_path=struct_filename_remote_path,
                                               local_path=struct_filename_local_path)
        # Списки из Экселя
        try:
            self.df_index_fea_code = pd.read_excel(self.index_filename, sheet_name='FEA_CODE')
            self.df_index_har_code1 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE1')

            # списки с описаниями в FT или Description
            har_code1_ft_list_RU = self.df_index_har_code1.loc[self.df_index_har_code1['LOC'] == 'FT']['RU'].tolist()
            har_code1_ft_list_EN = self.df_index_har_code1.loc[self.df_index_har_code1['LOC'] == 'FT']['EN'].tolist()
            har_code1_d_list_RU = self.df_index_har_code1.loc[self.df_index_har_code1['LOC'] == 'D']['RU'].tolist()
            har_code1_d_list_EN = self.df_index_har_code1.loc[self.df_index_har_code1['LOC'] == 'D']['EN'].tolist()

            self.df_index_har_code2 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE2')
            har_code2_ft_list_RU = self.df_index_har_code2.loc[self.df_index_har_code2['LOC'] == 'FT']['RU'].tolist()
            har_code2_ft_list_EN = self.df_index_har_code2.loc[self.df_index_har_code2['LOC'] == 'FT']['EN'].tolist()
            har_code2_d_list_RU = self.df_index_har_code2.loc[self.df_index_har_code2['LOC'] == 'D']['RU'].tolist()
            har_code2_d_list_EN = self.df_index_har_code2.loc[self.df_index_har_code2['LOC'] == 'D']['EN'].tolist()

            self.har_code_ft_list = har_code1_ft_list_RU + har_code2_ft_list_RU + har_code1_ft_list_EN + har_code2_ft_list_EN
            self.har_code_d_list = har_code1_d_list_RU + har_code2_d_list_RU + har_code1_d_list_EN + har_code2_d_list_EN

            self.df_index_fea_type = pd.read_excel(self.index_filename, sheet_name='FEA_TYPE')
            self.df_index_rep_method = pd.read_excel(self.index_filename, sheet_name='REP_METHOD')
        except OSError as e:
            if e.errno == errno.EACCES:
                print("# Permission ERROR: DBF_INDEX.xlsx")
                sys.exit()
        try:
            self.df_struct = pd.read_excel(self.struct_filename, sheet_name='STRUCT')
            self.df_struct_col_var = pd.read_excel(self.struct_filename, sheet_name='COL_VAR')
            self.df_struct_col_id_formats = pd.read_excel(self.struct_filename, sheet_name='COL_ID_STRUCT')
        except OSError as e:
            if e.errno == errno.EACCES:
                print("# Permission ERROR: STRUCT.xlsx")
                sys.exit()

        self.ml_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'ML', 'ID']
        self.geom_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'GEOM', 'ID']
        self.welds_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'WELD', 'ID']
        self.lam_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'LAM', 'ID']

        # FEA_RU
        # TYPE_RU
        # FEA_EN
        # TYPE_EN

    def load_dbf(self, dbf_path):

        self.dbf_path = dbf_path

        print("# DB load status: Loading...")

        dbf_raw = Dbf5(dbf_path, codec='cp1251')
        self.df_dbf = dbf_raw.to_dataframe()

        # переименовываем столбцы
        for i, row in self.df_struct.iterrows():
            struct_column_to_replace = str(self.df_struct.loc[i]["DBF_COL"])
            struct_column_target_name = str(self.df_struct.loc[i]["TARGET_COL_NAME"])
            if struct_column_to_replace in self.df_dbf.columns:
                self.df_dbf = self.df_dbf.rename(columns={struct_column_to_replace: struct_column_target_name})

        self.df_dbf['#FEA_CODE_REPLACE'] = self.df_dbf['#FEA_CODE']
        self.df_dbf['#HAR_CODE1_REPLACE'] = self.df_dbf['#HAR_CODE1']
        self.df_dbf['#HAR_CODE2_REPLACE'] = self.df_dbf['#HAR_CODE2']
        self.df_dbf['#KML_CLASS'] = self.df_dbf['#FEA_CODE']
        self.df_dbf['#FEATURE'] = self.df_dbf['#FEA_CODE']
        self.df_dbf['#FEA_CODE_REPLACE_FT'] = self.df_dbf['#FEA_CODE']
        self.df_dbf['#JL'] = ''
        self.df_dbf['#DIMM'] = ''
        self.df_dbf['#DESCR'] = ''
        # self.df_dbf['#REMAIN_WT'] = ''
        self.df_dbf['#ORIENT_HOUR'] = ''
        self.df_dbf['#BLANK'] = ''

        if "#CORR" not in self.df_dbf.columns:
            self.df_dbf['#CORR'] = ''

    def get_color_type_df(self, lng):
        return self.df_index_fea_code[['COLOR_TYPE', f'FEA_{lng}']]

    def get_ml_list(self):
        return self.ml_list

    def get_geom_list(self):
        return self.geom_list

    def get_welds_list(self):
        return self.welds_list

    def parse_columns(self, columns_list, ret_blank=True):

        """
        Получаем лист названий столбцов и возвращем соотвтствие их ID к Описанию
        ret_blank - если столбец не найден - возвращяем BLANK или оставляем искомое имя
        """

        col_id_list = []
        column_names = []

        if len(columns_list) > 0:
            for column_name in columns_list:
                cname_struct = '#BLANK'
                col_name = '#BLANK'
                col_id = '#BLANK'

                for i, row in self.df_struct_col_var.iterrows():
                    if column_name == self.df_struct_col_var.loc[i]['COL_VAR_NAME']:
                        cname_struct = self.df_struct_col_var.loc[i]['COL_ID']
                    elif cname_struct == '#BLANK':
                        cname_struct = column_name

                for i, row in self.df_struct_col_id_formats.iterrows():
                    if cname_struct == self.df_struct_col_id_formats.loc[i]['COL_ID']:
                        col_id = self.df_struct_col_id_formats.loc[i]['COL_ID']
                        col_name = self.df_struct_col_id_formats.loc[i][f'COL_NAME_{self.lng}']
                    elif col_id == '#BLANK':
                        if ret_blank:
                            col_name = '#BLANK'
                        else:
                            col_name = cname_struct

                col_id_list.append(col_id)
                column_names.append(col_name)

        return col_id_list, column_names

    def convert_dbf(self, diameter, dbf_path, lang):

        self.lng = lang

        self.load_dbf(dbf_path=dbf_path)

        if diameter < 100:
            diameter = diameter * 25.4

        self.df_replace(self.df_index_fea_code, '#FEA_CODE_REPLACE', change_class=True)

        self.df_replace(self.df_index_fea_code, '#FEATURE', fea=True)
        self.df_replace(self.df_index_fea_code, '#FEA_CODE_REPLACE_FT', ftype=True)

        self.df_replace(self.df_index_har_code1, '#HAR_CODE1_REPLACE')
        self.df_replace(self.df_index_har_code2, '#HAR_CODE2_REPLACE')
        self.df_replace(self.df_index_fea_type, '#LOC')
        self.df_replace(self.df_index_rep_method, '#REP_METHOD')

        # exportpath = r'd:\WORK\OrenburgNeft\NOA 8 inch DNS Olhovskaya to Terminal Service, 19.749 km\Reports\PR\Run3\DB_corr dist\123.csv'
        # self.df_dbf.to_csv(exportpath, encoding='cp1251', index=False)

        self.df_dbf['#FEATURE_TYPE'] = self.df_dbf.apply(
            lambda row: dbf_feature_type_combine(fea_code_replace_ft=row['#FEA_CODE_REPLACE_FT'],
                                                 har_code1=row['#HAR_CODE1_REPLACE'],
                                                 har_code2=row['#HAR_CODE2_REPLACE'], corr=row['#CORR'],
                                                 description=row['#DBF_DESCR'], ft_list=self.har_code_ft_list,
                                                 d_list=self.har_code_d_list), axis=1)

        self.df_dbf['#DESCR'] = self.df_dbf.apply(
            lambda row: dbf_description_combine(har_code1=row['#HAR_CODE1_REPLACE'],
                                                har_code2=row['#HAR_CODE2_REPLACE'], corr=row['#CORR'],
                                                description=row['#DBF_DESCR'], ft_list=self.har_code_ft_list,
                                                d_list=self.har_code_d_list), axis=1)

        # вычитаем и добавляем к дистации швов для корректной сортировки Аномалий и прочего на дистанции шва
        self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] = \
            self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] + 0.0001
        # сортировка по дистации
        self.df_dbf = self.df_dbf.sort_values('#DIST_START')
        self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] = \
            self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] - 0.0001

        # конвертирование Глубины в числа
        self.df_dbf['#DEPTH_MM'] = pd.to_numeric(self.df_dbf['#DEPTH_MM'], errors='coerce')

        # расчет остаточной толщины стенки
        # self.df_dbf['#REMAIN_WT'] = self.df_dbf.loc[self.df_dbf['#DEPTH_MM'] != '', '#WT'] - self.df_dbf.loc[
        #     self.df_dbf['#DEPTH_MM'] != '', '#DEPTH_MM']

        # конвертирование Номера особенности в числа
        self.df_dbf['#FEA_NUM'] = self.df_dbf['#FEA_NUM'].fillna(-1111111).astype(int)

        # https://stackoverflow.com/questions/25952790/convert-pandas-series-from-dtype-object-to-float-and-errors-to-nans

        # конвертирование WT в число
        self.df_dbf['#WT'] = pd.to_numeric(self.df_dbf['#WT'], errors='coerce')

        # конвертируем давление в MPa
        self.df_dbf['#PSAFE'] = self.df_dbf.loc[self.df_dbf['#PSAFE'] != '', '#PSAFE'] * 0.0980665

        # создание углов в градусах
        self.df_dbf['#ORIENT_HOUR'] = self.df_dbf.loc[self.df_dbf['#ORIENT_DEG'] != '', '#ORIENT_DEG'] / 720

        # глубины вмятина в процентах
        self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.geom_list), '#DEPTH_PRC'] = round(
            self.df_dbf['#DEPTH_MM'] / diameter * 100, 1)
        # глубины  процентах
        self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.ml_list) | self.df_dbf['#FEA_CODE'].isin(
            self.lam_list), '#DEPTH_PRC'] = round(self.df_dbf['#DEPTH_MM'] / self.df_dbf['#WT'] * 100, 1)
        # глубина '1' если получилась отрицательной (для проверки)
        self.df_dbf.loc[self.df_dbf['#DEPTH_PRC'] < 0, '#DEPTH_PRC'] = 1
        # чистка длин у швов
        self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#LENGTH'] = ''

        # нумеруем кастом JN
        self.df_dbf['#JN_Custom'] = ''
        for i, row in self.df_dbf.iterrows():
            if int(self.df_dbf.loc[i]['#FEA_CODE']) in self.welds_list.values:
                self.df_dbf.at[i, '#JN_Custom'] = (i + 1) * 10
        # меняем пустые на NAN и заполнем их предыдущими значениями
        self.df_dbf['#JN_Custom'] = self.df_dbf['#JN_Custom'].replace('', np.NAN)
        self.df_dbf['#JN_Custom'] = self.df_dbf['#JN_Custom'].fillna(method='ffill')
        # создаем фрэйм швов
        welds_jn_dist_array = self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list)][
            ['#JN_Custom', '#DIST_START', '#FEA_CODE']]
        # считаем длину секций
        welds_jn_dist_array['#JL'] = round(welds_jn_dist_array['#DIST_START'].diff(1), 3)
        # двигаем длину секций на 1 вверх
        welds_jn_dist_array['#JL'] = welds_jn_dist_array['#JL'].shift(-1)
        # заполняем длину секций общей базы по созданному фрэйму швов
        try:
            self.df_dbf['#JL'] = self.df_dbf['#JN_Custom'].map(welds_jn_dist_array.set_index('#JN_Custom')['#JL'])
        except Exception as ex:
            print('## ERROR: no Welds, JL not calculated: ', ex)

        # маска для потерей металла под расчет DIMM
        mask = self.df_dbf['#FEA_CODE'].isin(self.ml_list) & \
               self.df_dbf["#LENGTH"].notnull() & \
               self.df_dbf["#WIDTH"].notnull() & \
               self.df_dbf["#WT"].notnull()

        try:
            self.df_dbf['#DIMM'] = self.df_dbf[mask].apply(
                lambda row: dimm(length=row['#LENGTH'], width=row['#WIDTH'], wt=row['#WT'],
                                 return_format=self.lng), axis=1)
        except Exception as ex:
            print('# Info: no ML in DB detected')

        self.df_dbf = self.df_dbf.replace({'nan': '', 'NaN': '', float('NaN'): '', -1111111: ''})

        print("# DB load status: Loaded successfully!\n")

        return self.df_dbf

    def df_replace(self, df_what, replace_column_name, change_class=False, fea=False, ftype=False):

        for i, row in df_what.iterrows():
            if i < len(df_what):
                f_ID = int(df_what.loc[i]['ID'])

                if fea is False and ftype is False:
                    f_ID_DSCR = str(df_what.loc[i][self.lng])
                if fea is True:
                    f_ID_DSCR = str(df_what.loc[i][f'FEA_{self.lng}'])
                if ftype is True:
                    f_ID_DSCR = str(df_what.loc[i][f'TYPE_{self.lng}'])
                # print(i, " - ", f_ID, " - ", f_ID_DSCR)
                self.df_dbf.loc[self.df_dbf[replace_column_name] == f_ID, replace_column_name] = f_ID_DSCR
                if change_class is True:
                    f_ID_CLASS = str(df_what.loc[i]['KML_CLASS'])
                    self.df_dbf.loc[self.df_dbf['#KML_CLASS'] == f_ID, '#KML_CLASS'] = f_ID_CLASS
        # print("Replace ", replace_column_name, " done")

    def fea_type_parse(self):

        id_list = self.df_index_fea_code['ID']

        for i, row in self.df_dbf.iterrows():
            ind = self.df_dbf.loc[i]['#FEA_CODE']
            if not id_list[id_list.isin([ind])].empty:
                ID_ROW = self.df_index_fea_code.loc[self.df_index_fea_code['ID'] == ind]
                self.df_dbf.loc[i]['#FEATURE'] = ID_ROW.loc[1]['FEA_RU']
                self.df_dbf.loc[i]['#FEATURE_TYPE'] = ID_ROW.loc[1]['TYPE_RU']


# сохраняем в csv c custom столбцами
def to_csv_custom_header(df, csv_path, column_names, csv_encoding):
    print(f'# Custom headers info: Total columns: {len(column_names)}')
    # print(f'# Custom headers info: Columns list: {column_names}')

    with open(csv_path, 'w') as csvfile:
        column_names_string = ''
        for i in column_names:
            column_names_string = column_names_string + f'"{i}",'
        # column_names_string = ",".join([str(i) for i in column_names])
        csvfile.write(f'{column_names_string}\n')

        # write = csv.writer(csvfile)
        # write.writerow(column_names)
        # csvfile.write('КонЭц\n')

    df.to_csv(csv_path, header=False, index=False, encoding=csv_encoding, mode="a")


def cross_columns_list(list_exist, list_target):
    """
    Возращаем список который входит в Exist
    cross_columns = cross_columns_list(total_columns, exp_format)

    list_exist - столбцы в наличии
    list_target - желаемый список столбцов

    cross_columns - только те из list_target что имеются в exist
    """

    cross_columns = []
    for val in list_target:
        for val2 in list_exist:
            if val == val2:
                cross_columns.append(val)
    return cross_columns


if __name__ == '__main__':

    DEBUG = 0

    # дэфолтные
    diam = None
    lang = "RU"
    custom_columns = []

    if DEBUG == 1:
        path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\Test\3nocm.DBF"
    else:
        arg = ''
        for arg in sys.argv[1:]:
            print(arg)

        if arg != '':
            path = arg
        else:
            path = input("DBF path: ")

            custom_columns = input("Enter columns: ")
            custom_columns = custom_columns.split('\t')

            print(f'# ARG columns info: Columns received: {len(custom_columns)}')
            print(f'# ARG columns info: Columns list: {custom_columns}')

            # print(custom_columns)

        if path[-3:] != 'DBF' and path[-3:] != 'dbf':
            input("##ERROR: not DBF link")
        else:
            # проверяем на наличие дюйма по проекту
            diam = parse_inch_prj(path)
            # если не нашли то ищем на всякий в проекте в папке выше (если под папка рабочих)
            if diam is None:
                diam = parse_inch_prj(os.path.dirname(path))

            # если не нашли - просим ввести дюймаж
            if diam is None:
                # просим ввести дюймаж
                diam = input('Enter Diameter (Inch/mm)?: ')
                # если что-то ввели то поверяем на дюйм/мм и домножаем
                if diam != '':
                    diam = float(diam)
                    if diam > 100:
                        diam = diam / 25.4
                # если ввели пустоту
                else:
                    # ищем дюймаж в пути файла
                    diam = parse_inch_path(path)
                    # если и в пути нет, то пишем дефолтный
                    if diam is None:
                        print("# Default diameter = 8.625 inch")
                        diam = 8.625

            if arg == '':
                lang = input("'1/2' for EN/SP?: ")
            else:
                lang = "EN"

            if lang == '1':
                lang = "EN"
            elif lang == '2':
                lang = "SP"
            else:
                lang = "RU"

    df_dbf = df_DBF(DBF_path=path, lang=lang)
    exp = df_dbf.convert_dbf(diameter=diam)

    # возвращаем столбец что нашли и их перевод для Кастом
    custom_columns, custom_columns_names = df_dbf.parse_columns(columns_list=custom_columns)

    # Thailand
    # exp_format1 = ['#FEA_NUM', '#DIST_START', '#JN', '#US', '#JL', '#FEATURE', '#FEATURE_TYPE', '#DIMM', '#ORIENT_HOUR',
    #               '#WT', '#LENGTH', '#WIDTH', '#DEPTH_PRC', '#DEPTH_MM', '#RWT', '#LOC', '#ERF', '#PSAFE',
    #               '#CLUSTER', '#DESCR', '#LAT', '#LONG', '#ALT']

    # , '#FEA_CODE_REPLACE','#HAR_CODE1_REPLACE', '#HAR_CODE2_REPLACE', '#DBF_DESCR', '#REMARKS'

    # exp_format2 = ['#JN', '#FEA_NUM', '#DIST_START', '#JL', '#US', '#DOC', '#FEA_CODE_REPLACE', '#HAR_CODE1_REPLACE',
    #               '#HAR_CODE2_REPLACE', '#FEATURE_TYPE', '#DBF_DESCR', "#DESCR", '#CORR', '#ERF', '#CLUSTER',
    #               '#PSAFE', '#WT', '#DEPTH_MM', '#DEPTH_PRC', '#AMPL', '#DEPTH_PREV',
    #               '#LENGTH', '#WIDTH', '#ORIENT_DEG', '#LOC', '#DIMM', '#LAT', '#LONG', '#ALT']

    # orenburg

    total_columns = exp.columns.values.tolist()

    # возвращаем пересечение от шаблона к имеющимся
    cross_columns = cross_columns_list(total_columns, exp_format)
    # возвращаем столбец что нашли и их перевод для Дэфолтного
    cross_columns_return, column_names_cross = df_dbf.parse_columns(columns_list=cross_columns)

    # добавляем ДОК столбцы в кастом столбцы
    custom_columns.append('#DOC')
    custom_columns_names.append('#DOC')

    # если есть кастом, то экспортим его, в противном случае - шаблон
    if len(custom_columns) > 1:
        exp1 = exp[custom_columns]
        column_names = custom_columns_names
    else:
        exp1 = exp[cross_columns]
        column_names = column_names_cross

    # with open(file_path, 'w') as f:
    #    f.write('Custom String\n')

    # df.to_csv(file_path, header=False, mode="a")
    # print (cross_columns)

    absbath = os.path.dirname(path)
    basename = os.path.basename(path)
    exportpath = os.path.join(absbath, basename)
    exportpath_csv = f'{exportpath[:-4]}.csv'
    exportpath_xlsx = f'{exportpath[:-4]}.xlsx'

    if lang == "RU":
        csv_encoding = "cp1251"
    else:
        csv_encoding = "utf-8"
    try:
        to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names, csv_encoding=csv_encoding)
        # exp1.to_csv(exportpath_csv, encoding=csv_encoding, index=False)
        # xw.Book(exportpath_csv)
        # xw.view(exportpath_csv, table=False)
        if lang == "SP":
            exp1.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)

    except Exception as PermissionError:
        print(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
        exportpath = os.path.join(absbath, basename)
        exportpath_csv = f'{exportpath[:-4]}_1.csv'
        exportpath_xlsx = f'{exportpath[:-4]}_1.xlsx'
        to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names, csv_encoding=csv_encoding)
        # exp1.to_csv(exportpath_csv, encoding=csv_encoding, index=False)
        if lang == "SP":
            exp1.to_excel(exportpath_xlsx, encoding=csv_encoding, index=False)

    if DEBUG != 1:
        input("~~~ Done~~~")
    else:
        print("~~~ DEBUG Done ~~~")
