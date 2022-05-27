import os.path

import pandas as pd
import numpy as np
from simpledbf import Dbf5


class df_DBF:
    def __init__(self, DBF_path, lang="RU"):

        self.lng = lang
        self.dbf_path = DBF_path
        self.index_filename = 'DBF_INDEX.xlsx'

        self.df_index_fea_code = pd.read_excel(self.index_filename, sheet_name='FEA_CODE')
        self.df_index_har_code1 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE1')
        self.df_index_har_code2 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE2')
        self.df_index_fea_type = pd.read_excel(self.index_filename, sheet_name='FEA_TYPE')

        dbf_raw = Dbf5(self.dbf_path, codec='cp1251')
        self.df_dbf = dbf_raw.to_dataframe()
        self.df_dbf['Feature'] = self.df_dbf['FEA_CODE']
        self.df_dbf['CLASS'] = self.df_dbf['FEA_CODE']

    def convert_dbf(self):

        self.df_replace(self.df_index_fea_code, "Feature", change_class=True)
        self.df_replace(self.df_index_har_code1, "HAR_CODE1")
        self.df_replace(self.df_index_har_code2, "HAR_CODE2")
        self.df_replace(self.df_index_fea_type, "FEA_TYPE")

        self.df_dbf = self.df_dbf.sort_values('FEA_DIST')
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
    #path = input("enter DBF path: ")

    path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\1nldm.dbf"

    df_dbf = df_DBF(DBF_path=path)
    exp = df_dbf.convert_dbf()

    exp['FEA_DEPTH'] = pd.to_numeric(exp['FEA_DEPTH'], errors='coerce')

    # https://stackoverflow.com/questions/25952790/convert-pandas-series-from-dtype-object-to-float-and-errors-to-nans

    exp['WT'] = pd.to_numeric(exp['WT'], errors='coerce')

    #exp.loc[exp['FEA_CODE'] == 201, 'COMMENT'] = "Геом"
    #exp.loc[exp['FEA_CODE'] == 101, 'COMMENT'] = "ЭмЭль"

    exp['FEA_DEPTH_M'] = exp['FEA_DEPTH'] / exp['WT'] * 100

    exp['FEA_DEPTH_M'] = exp['FEA_DEPTH'] / exp['WT'] * 100

    exp1 = exp[
        ["FEA_NUM", "FEA_DIST", "DOC", "Feature", "HAR_CODE1", "HAR_CODE2", "COMMENT", "FEA_TYPE", "FEA_DEPTH",
         'FEA_DEPTH_M', "WT",
         "LATITUDE", "LONGITUDE", "CLASS"]]

    absbath = os.path.dirname(path)
    basename = os.path.basename(path)
    exportpath = os.path.join(absbath, basename)

    print(exportpath)
    exp1.to_csv(f'{exportpath}.csv', encoding='cp1251', index=False)
