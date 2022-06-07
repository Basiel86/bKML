import os.path
import pandas as pd
import sys
from simpledbf import Dbf5


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

        self.df_index_fea_code = pd.read_excel(self.index_filename, sheet_name='FEA_CODE')
        self.df_index_har_code1 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE1')
        self.df_index_har_code2 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE2')
        self.df_index_fea_type = pd.read_excel(self.index_filename, sheet_name='FEA_TYPE')

        self.ml_list = [101, 106, 32869, 3581198, 1081016421]
        self.geom_list = [201, 32969, 52887753]

        # FEA_RU
        # TYPE_RU
        # FEA_EN
        # TYPE_EN

        dbf_raw = Dbf5(self.dbf_path, codec='cp1251')
        self.df_dbf = dbf_raw.to_dataframe()
        self.df_dbf['Feature'] = self.df_dbf['FEA_CODE']
        self.df_dbf['CLASS'] = self.df_dbf['FEA_CODE']

    def convert_dbf(self, diameter):

        self.df_replace(self.df_index_fea_code, "Feature", change_class=True)
        self.df_replace(self.df_index_har_code1, "HAR_CODE1")
        self.df_replace(self.df_index_har_code2, "HAR_CODE2")
        self.df_replace(self.df_index_fea_type, "FEA_TYPE")

        self.df_dbf = self.df_dbf.sort_values('FEA_DIST')

        self.df_dbf['FEA_DEPTH'] = pd.to_numeric(self.df_dbf['FEA_DEPTH'], errors='coerce')

        # https://stackoverflow.com/questions/25952790/convert-pandas-series-from-dtype-object-to-float-and-errors-to-nans

        self.df_dbf['WT'] = pd.to_numeric(self.df_dbf['WT'], errors='coerce')

        self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.geom_list), 'FEA_DEPTH_PRC'] = round(
            self.df_dbf['FEA_DEPTH'] / diameter * 100, 1)
        self.df_dbf.loc[self.df_dbf['FEA_CODE'].isin(self.ml_list), 'FEA_DEPTH_PRC'] = round(
            self.df_dbf['FEA_DEPTH'] / self.df_dbf['WT'] * 100, 1)
        self.df_dbf.loc[self.df_dbf['FEA_DEPTH_PRC'] < 0, 'FEA_DEPTH_PRC'] = 1

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


if __name__ == '__main__':
    #path = input("DBF path: ")
    #diam = input("INCH?: ")
    # lang = input("'1' for EN?: ")

    path = r"WORK_DBF\1nhdm.DBF"
    diam = 10.75
    lang = 'RU'
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

    df_dbf = df_DBF(DBF_path=path, lang = lang)
    exp = df_dbf.convert_dbf(diameter=diam)

    exp_format = ['SECT_NUM', 'FEA_NUM', 'FEA_DIST', 'REL_DIST', 'DOC', 'Feature', 'HAR_CODE1', 'HAR_CODE2', 'COMMENT',
                  'REMARKS', 'ERF', 'CLUST_NUM',
                  'MAOP', 'WT', 'FEA_DEPTH', 'FEA_DEPTH_PRC', 'AMPL_MFL', 'DEPTH_PREV', 'DEPTH_TMP', 'DEPTH_TMP2',
                  'FEA_LENGTH', 'FEA_WIDTH', 'CLCK_DP', 'FEA_TYPE', 'REP_GROUP', 'LATITUDE', 'LONGITUDE', 'HEIGHT']

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
    exp1.to_csv(exportpath, encoding='cp1251', index=False)
    print(exportpath)
    input("~~~ Done~~~")
