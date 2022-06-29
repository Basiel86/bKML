import os.path
import pandas as pd
import sys
from simpledbf import Dbf5
from parse_inch import parse_inch
from dimm import dimm

"""
    FEA_CODE_REPLACE - столбец прямой конверсии

"""


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class df_DBF:
    def __init__(self, DBF_path, lang="RU"):

        self.lng = lang
        self.dbf_path = DBF_path
        self.index_filename = resource_path('DBF_INDEX.xlsx')

        # Списки из Экселя
        self.df_index_fea_code = pd.read_excel(self.index_filename, sheet_name='FEA_CODE')
        self.df_index_har_code1 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE1')
        self.df_index_har_code2 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE2')
        self.df_index_fea_type = pd.read_excel(self.index_filename, sheet_name='FEA_TYPE')

        self.ml_list = [101, 106, 32869, 3581198, 1081016421]
        self.geom_list = [201, 32969, 52887753]
        self.welds_list = [901, 902, 903, 904, 905, 906]

        # FEA_RU
        # TYPE_RU
        # FEA_EN
        # TYPE_EN

        dbf_raw = Dbf5(self.dbf_path, codec='cp1251')
        self.df_dbf = dbf_raw.to_dataframe()
        self.df_dbf['FEA_CODE_REPLACE'] = self.df_dbf['FEA_CODE']
        self.df_dbf['CLASS'] = self.df_dbf['FEA_CODE']
        self.df_dbf['FEATURE'] = ''
        self.df_dbf['FEATURE_TYPE'] = ''
        self.df_dbf['JL'] = ''
        self.df_dbf['DIMM'] = ''

    def convert_dbf(self, diameter):

        if diameter < 100:
            diameter = diameter * 25.4

        self.df_replace(self.df_index_fea_code, "FEA_CODE_REPLACE", change_class=True)
        self.df_replace(self.df_index_har_code1, "HAR_CODE1")
        self.df_replace(self.df_index_har_code2, "HAR_CODE2")
        self.df_replace(self.df_index_fea_type, "FEA_TYPE")
        # self.fea_type_parse()

        # сортировка по дистации
        self.df_dbf = self.df_dbf.sort_values('FEA_DIST')
        # конвертирование Глубины в числа
        self.df_dbf['FEA_DEPTH'] = pd.to_numeric(self.df_dbf['FEA_DEPTH'], errors='coerce')

        # https://stackoverflow.com/questions/25952790/convert-pandas-series-from-dtype-object-to-float-and-errors-to-nans

        # конвертирование WT в число
        self.df_dbf['WT'] = pd.to_numeric(self.df_dbf['WT'], errors='coerce')

        # глубины вмятина в процентах
        self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.geom_list), 'FEA_DEPTH_PRC'] = round(
            self.df_dbf['FEA_DEPTH'] / diameter * 100, 1)
        # глубины потерь в процентах
        self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.ml_list), 'FEA_DEPTH_PRC'] = round(
            self.df_dbf['FEA_DEPTH'] / self.df_dbf['WT'] * 100, 1)
        # глубина '1' если получилась отрицательной (для проверки)
        self.df_dbf.loc[self.df_dbf['FEA_DEPTH_PRC'] < 0, 'FEA_DEPTH_PRC'] = 1
        # чистка длин у швов
        self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.welds_list), 'FEA_LENGTH'] = ''

        # создаем фрэйм швов
        welds_jn_dist_array = self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.welds_list)][['SECT_NUM', 'FEA_DIST']]
        # считаем длину секций
        welds_jn_dist_array['JL'] = round(welds_jn_dist_array['FEA_DIST'].diff(1), 3)
        # двигаем длину секций на 1 вверх
        welds_jn_dist_array['JL'] = welds_jn_dist_array['JL'].shift(-1)
        # заполняем длину секций общей базы по созданному фрэйму швов
        self.df_dbf['JL'] = self.df_dbf['SECT_NUM'].map(welds_jn_dist_array.set_index('SECT_NUM')['JL'])

        mask = self.df_dbf['FEA_CODE'].isin(self.ml_list) & \
               self.df_dbf.FEA_LENGTH.notnull() & \
               self.df_dbf.FEA_WIDTH.notnull() & \
               self.df_dbf.WT.notnull()

        self.df_dbf['DIMM'] = self.df_dbf[mask].apply(
            lambda row: dimm(length=row['FEA_LENGTH'], width=row['FEA_WIDTH'], wt=row['WT'], return_format=self.lng), axis=1)

        # print(self.df_dbf['DIMM'])

        self.df_dbf = self.df_dbf.replace({'nan': '', 'NaN': '', float('NaN'): ''})

        return self.df_dbf

    def df_replace(self, df_what, replace_column_name, change_class=False):

        for i, row in df_what.iterrows():
            if i < len(df_what):
                f_ID = int(df_what.loc[i]['ID'])
                f_ID_DSCR = str(df_what.loc[i][self.lng])
                # print(i, " - ", f_ID, " - ", f_ID_DSCR)
                self.df_dbf.loc[self.df_dbf[replace_column_name] == f_ID, replace_column_name] = f_ID_DSCR
                if change_class is True:
                    f_ID_CLASS = str(df_what.loc[i]['CLASS'])
                    self.df_dbf.loc[self.df_dbf['CLASS'] == f_ID, 'CLASS'] = f_ID_CLASS
        # print("Replace ", replace_column_name, " done")

    def fea_type_parse(self):

        id_list = self.df_index_fea_code['ID']

        for i, row in self.df_dbf.iterrows():
            ind = self.df_dbf.loc[i]['FEA_CODE']
            if not id_list[id_list.isin([ind])].empty:
                ID_ROW = self.df_index_fea_code.loc[self.df_index_fea_code['ID'] == ind]
                self.df_dbf.loc[i]['FEATURE'] = ID_ROW.loc[1]['FEA_RU']
                self.df_dbf.loc[i]['FEATURE_TYPE'] = ID_ROW.loc[1]['TYPE_RU']


if __name__ == '__main__':

    DEBUG = 0

    if DEBUG == 1:
        path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\Test\1nxbu.dbf"
        path1 = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\Test\2nsiu.DBF"
        lang = "RU"
        diam = 254
    else:
        arg = ''
        for arg in sys.argv[1:]:
            print(arg)

        if arg != '':
            path = arg
        else:
            path = input("DBF path: ")

        if path[-3:] != 'DBF' and path[-3:] != 'dbf':
            input("Error: not DBF link")
        else:
            diam = parse_inch(path)
            if diam is not None:
                print('Detected Diam: ', diam, " inch")
            else:
                diam = input("INCH?: ")

            if arg == '':
                lang = input("'1' for EN?: ")
            else:
                lang = "RU"

            # path = r"WORK_DBF\1nhdm.DBF"
            # diam = 10.75
            # lang = 'RU'
            if diam == '':
                diam = 1
            else:
                diam = float(diam)

            if diam < 100:
                diam = diam * 25.4

            if lang == '1':
                lang = "EN"
            else:
                lang = "RU"

    print('\nProcess...\n')
    df_dbf = df_DBF(DBF_path=path, lang=lang)
    exp = df_dbf.convert_dbf(diameter=diam)

    exp_format = ['SECT_NUM', 'FEA_NUM', 'FEA_DIST', 'JL', 'REL_DIST', 'DOC', 'FEA_CODE_REPLACE', 'HAR_CODE1',
                  'HAR_CODE2',
                  'COMMENT',
                  'REMARKS', 'ERF', 'CLUST_NUM',
                  'MAOP', 'WT', 'FEA_DEPTH', 'FEA_DEPTH_PRC', 'AMPL_MFL', 'DEPTH_PREV', 'DEPTH_TMP', 'DEPTH_TMP2',
                  'FEA_LENGTH', 'FEA_WIDTH', 'CLCK_DP', 'FEA_TYPE', 'DIMM', 'LATITUDE', 'LONGITUDE', 'HEIGHT']

    total_coluns = exp.columns.values.tolist()
    cross_columns = []
    for val in exp_format:
        for val2 in total_coluns:
            if val == val2:
                cross_columns.append(val)

    # print(cross_columns)

    exp1 = exp[cross_columns]

    absbath = os.path.dirname(path)
    basename = os.path.basename(path)
    exportpath = os.path.join(absbath, basename)
    exportpath = f'{exportpath[:-4]}.csv'
    try:
        exp1.to_csv(exportpath, encoding='cp1251', index=False)
    except Exception as PermissionError:
        print(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
        exportpath = os.path.join(absbath, basename)
        exportpath = f'{exportpath[:-4]}_1.csv'
        exp1.to_csv(exportpath, encoding='cp1251', index=False)

    if DEBUG != 1:
        input("~~~ Done~~~")
    else:
        print("~~~ Done~~~")
