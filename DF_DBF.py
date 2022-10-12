import os.path

import pandas as pd
import numpy as np
import sys
from simpledbf import Dbf5
from parse_inch import parse_inch_prj, parse_inch_combine
from dimm import dimm
import errno
from other_functions import dbf_feature_type_combine
from other_functions import dbf_description_combine
from parse_templates import read_template
import traceback
import codecs
import shutil
import time
import logging

# столбцы для дэфолтного экспорта
from Export_columns import exp_format

logger = logging.getLogger('app.DF_DBF')

# index_filename_remote_path = r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\DBF_INDEX.xlsx'
index_filename_local_path = r'IDs\DBF_INDEX.xlsx'
# struct_filename_remote_path = r'\\vasilypc\Vasily Shared (Read Only)\_Templates\PT\IDs\STRUCT.xlsx'
struct_filename_local_path = r'IDs\STRUCT.xlsx'


def print_time_spent(time_start, time_end, descr=""):
    print(f"### Time spent {descr}: {round(time_end - time_start, 3)}")


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


# Берем версию с компа / удаленно / инишника
def remote_or_local(remote_path, local_path, cfg_path, local):
    filename = os.path.basename(local_path)[:-5]

    if local is True:
        print(f"# Info: {filename} file info: Local (INI)")
        logger.debug(f"#{filename} file info: Local (INI)")
        return local_path

    try:
        if cfg_path is not None:
            if os.path.exists(cfg_path):
                print(f"# Info: {filename} file info: INI")
                logger.debug(f"{filename} file info: INI")
                return cfg_path

        if remote_path is not None:
            if os.path.exists(remote_path):
                print(f"# Info: {filename} file info: Remote")
                logger.debug(f"{filename} file info: Remote")
                return remote_path

        print(f"# Info: {filename} file info: Local")
        logger.debug(f"{filename} file info: Local")
        return resource_path(local_path)
    except Exception as ex:
        input(f"### ERROR: Remote or Local proc: {ex}")
        logger.exception('Remote or Local processing')


class df_DBF():
    def __init__(self, index_remote_path=None, struct_remote_path=None, index_path=None, struct_path=None, local=False):

        self.dbf_path = ''
        self.lng = 'RU'

        self.index_filename = remote_or_local(remote_path=index_remote_path,
                                              local_path=index_filename_local_path,
                                              cfg_path=index_path,
                                              local=local)
        self.struct_filename = remote_or_local(remote_path=struct_remote_path,
                                               local_path=struct_filename_local_path,
                                               cfg_path=struct_path,
                                               local=local)
        # Списки из Экселя
        try:
            self.df_index_fea_code = pd.read_excel(self.index_filename, sheet_name='FEA_CODE')
            self.df_index_har_code1 = pd.read_excel(self.index_filename, sheet_name='HAR_CODE1')

            self.df_dbf = None

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
                logger.error('Permission ERROR: DBF_INDEX.xlsx')
                sys.exit()
        try:
            self.df_struct = pd.read_excel(self.struct_filename, sheet_name='STRUCT')
            self.df_struct_col_var = pd.read_excel(self.struct_filename, sheet_name='COL_VAR')
            self.df_struct_col_id_formats = pd.read_excel(self.struct_filename, sheet_name='COL_ID_STRUCT')
        except OSError as e:
            if e.errno == errno.EACCES:
                print("# Permission ERROR: STRUCT.xlsx")
                logger.error('Permission ERROR: STRUCT.xlsx')
                sys.exit()

        self.ml_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'ML', 'ID']
        self.geom_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'GEOM', 'ID']
        self.welds_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'WELD', 'ID']
        self.lam_list = self.df_index_fea_code.loc[self.df_index_fea_code['KML_D_TYPE'] == 'LAM', 'ID']

    def load_dbf(self, dbf_path):

        self.dbf_path = dbf_path

        try:
            print("# DB load status: Loading...")

            logger.info(msg='#'*len(self.dbf_path))
            logger.debug(f'Get DBF file: {self.dbf_path}')
            dbf_raw = Dbf5(dbf_path, codec='cp1251')
            self.df_dbf = dbf_raw.to_dataframe()

            if 'F' not in self.df_dbf.columns:
                print('### WARNING: DB has not Proper Format, please run DBColumns First!!!')
                logger.warning('### WARNING: DB has not Proper Format, please run DBColumns First!!!')

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

            if "#REMARKS" not in self.df_dbf.columns:
                self.df_dbf['#REMARKS'] = ''

            if "#WT" not in self.df_dbf.columns:
                self.df_dbf['#WT'] = ''

            if "F" not in self.df_dbf.columns:
                self.df_dbf['F'] = ''

        except Exception as ex:
            print(f'### ERROR: Unhandled load_dbf error: {ex}')
            logger.exception('Unhandled load_dbf error')

    def get_color_type_df(self, lng):
        return self.df_index_fea_code[['COLOR_TYPE', f'FEA_{lng}']]

    def get_ml_list(self):
        return self.ml_list

    def get_geom_list(self):
        return self.geom_list

    def get_welds_list(self):
        return self.welds_list

    def parse_columns_df(self, columns_list: list, cross_list: list, ret_blank=True) -> pd.DataFrame:

        """
        Получаем лист названий столбцов и возвращем соотвтствие их ID к Описанию
        ret_blank - если столбец не найден - возвращяем BLANK или оставляем искомое имя
        cross_list - список столбцов для поиска пересечения
        Возвращаем DF со столбцами [COL_VAR_NAME, COL_ID, COL_NAME]
        """

        # Конвертим лист в Датафрейм и добавляем нужные столбцы
        columns_list_df = pd.DataFrame(columns_list, columns=["COL_VAR_NAME"])
        columns_list_df['COL_ID'] = ''
        columns_list_df['COL_NAME'] = ''

        # Конвертим Cross лист в Датафрейм и добавляем нужные столбцы
        cross_list_df = pd.DataFrame(cross_list, columns=["COL_ID"])
        cross_list_df = cross_list_df.set_index('COL_ID')

        # VAR структуру фильтруем по наличию в исходной
        filtered_var_id = self.df_struct_col_var[
            self.df_struct_col_var['COL_VAR_NAME'].isin(columns_list_df['COL_VAR_NAME'])]
        # Индексируем по столбцу COL_VAR_NAME
        filtered_var_id = filtered_var_id.set_index('COL_VAR_NAME')

        # Индексируем ID/NAME список по COL_ID
        col_id_name_df = self.df_struct_col_id_formats.set_index('COL_ID')

        # Бежим по полученному списку
        for i, row in columns_list_df.iterrows():
            # текущее имя VAR
            var_name = columns_list_df.loc[i]['COL_VAR_NAME']
            if var_name in filtered_var_id.index:
                # найденное ID по VAR
                var_id = filtered_var_id.loc[var_name]['COL_ID']

                # пишем ID и Name
                columns_list_df.loc[i]['COL_ID'] = var_id
                columns_list_df.loc[i]['COL_NAME'] = col_id_name_df.loc[var_id][f'COL_NAME_{self.lng}']
            else:
                columns_list_df.loc[i]['COL_ID'] = '#BLANK'
                if ret_blank:
                    col_name = '#BLANK'
                else:
                    col_name = var_name
                columns_list_df.loc[i]['COL_NAME'] = col_name

        # проверяем на наличие пересечения для получения #NOT_IN_DB
        for i, row in columns_list_df.iterrows():
            var_id = columns_list_df.loc[i]['COL_ID']
            if var_id not in cross_list_df.index and ret_blank:
                columns_list_df.loc[i][f'COL_ID'] = '#BLANK'
                columns_list_df.loc[i][f'COL_NAME'] = '#NOT_IN_DB'

        return columns_list_df

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

        ts = time.time()
        self.load_dbf(dbf_path=dbf_path)
        logger.debug(f'Timer: DBF load time: {round(time.time()-ts,3)} s.')
        # print_time_spent(ts, time.time(), descr="'DBF Load'")

        if diameter < 100:
            diameter = diameter * 25.4

        tsr = time.time()
        self.df_replace(self.df_index_fea_code, '#FEA_CODE_REPLACE', change_class=True)
        self.df_replace(self.df_index_fea_code, '#FEATURE', fea=True)
        self.df_replace(self.df_index_fea_code, '#FEA_CODE_REPLACE_FT', ftype=True)
        self.df_replace(self.df_index_har_code1, '#HAR_CODE1_REPLACE')
        self.df_replace(self.df_index_har_code2, '#HAR_CODE2_REPLACE')
        self.df_replace(self.df_index_fea_type, '#LOC')
        self.df_replace(self.df_index_rep_method, '#REP_METHOD')

        # print_time_spent(tsr, time.time(), descr="'REPLACE'")
        logger.debug(f'Timer: Replace time: {round(time.time()-tsr,3)} s.')
        ts = time.time()

        # exportpath = r'd:\WORK\OrenburgNeft\NOA 8 inch DNS Olhovskaya to Terminal Service, 19.749 km\Reports\PR\Run3\DB_corr dist\123.csv'
        # self.df_dbf.to_csv(exportpath, encoding='cp1251', index=False)

        try:
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
                self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] - 0.0001
            # сортировка по дистации
            self.df_dbf = self.df_dbf.sort_values('#DIST_START')
            self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] = \
                self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.welds_list), '#DIST_START'] + 0.0001

            # конвертирование Глубины в числа
            self.df_dbf['#DEPTH_MM'] = pd.to_numeric(self.df_dbf['#DEPTH_MM'], errors='coerce')

            # расчет остаточной толщины стенки
            # self.df_dbf['#REMAIN_WT'] = self.df_dbf.loc[self.df_dbf['#DEPTH_MM'] != '', '#WT'] - self.df_dbf.loc[
            #     self.df_dbf['#DEPTH_MM'] != '', '#DEPTH_MM']

            # конвертирование Номера особенности в числа
            self.df_dbf['#FEA_NUM'] = self.df_dbf['#FEA_NUM'].fillna(-1111111).astype(int)

            # https://stackoverflow.com/questions/25952790/convert-pandas-series-from-dtype-object-to-float-and-errors-to-nans

            # конвертируем давление в MPa
            if '#PSAFE' in self.df_dbf.columns:
                self.df_dbf['#PSAFE'] = self.df_dbf.loc[self.df_dbf['#PSAFE'] != '', '#PSAFE'] * 0.0980665

            # создание углов в градусах
            if '#ORIENT_HOUR' in self.df_dbf.columns:
                self.df_dbf['#ORIENT_HOUR'] = self.df_dbf.loc[self.df_dbf['#ORIENT_DEG'] != '', '#ORIENT_DEG'] / 720

            # глубины вмятина в процентах
            self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.geom_list), '#DEPTH_PRC'] = round(
                self.df_dbf['#DEPTH_MM'] / diameter * 100, 2)

            # конвертирование WT в число
            self.df_dbf['#WT'] = pd.to_numeric(self.df_dbf['#WT'], errors='coerce')

            # глубины  процентах
            self.df_dbf.loc[self.df_dbf['#FEA_CODE'].isin(self.ml_list) | self.df_dbf['#FEA_CODE'].isin(
                self.lam_list), '#DEPTH_PRC'] = round(self.df_dbf['#DEPTH_MM'] / self.df_dbf['#WT'] * 100, 2)
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
                logger.error('no Welds, JL not calculated')

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
                logger.error('no ML in DB detected')

            self.df_dbf = self.df_dbf.replace({'nan': '', 'NaN': '', float('NaN'): '', -1111111: ''})

            print("# DB load status: Loaded successfully!\n")
            logger.info("DB load status: Loaded successfully")

            # print_time_spent(ts, time.time(), descr="'DBF Convert to DF'")
            logger.debug(f'Timer: DBF to DF conversion: {round(time.time()-ts,3)} s.')

            return self.df_dbf

        except ValueError:
            print(traceback.format_exc())
            input("\n### Error: DB Error, try process it by DBFnewColumns and repeat, if So, call Developer!")
            logger.exception('DB Error, try process it by DBFnewColumns and repeat, if So, call Developer!')
        except Exception as ex:
            print(ex)
            input(traceback.format_exc())
            logger.exception("Something went wrong")

    def df_replace(self, df_what, replace_column_name, change_class=False, fea=False, ftype=False):

        if replace_column_name in self.df_dbf.columns:
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


def export_default(dbf_path):

    logger.info("DB Export")

    lang = "RU"
    path = dbf_path

    print(f'### Fast Export path: {dbf_path}')

    diam = parse_inch_prj(path)

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
            diam = parse_inch_prj(path)
            # если и в пути нет, то пишем дефолтный
            if diam is None:
                print("# Default diameter = 8.625 inch")
                diam = 8.625

    df_dbf = df_DBF()
    exp = df_dbf.convert_dbf(dbf_path=path, lang=lang, diameter=diam)

    total_columns = exp.columns.values.tolist()
    default_columns_list = read_template("Default")

    # получаем пересечение Шаблона с имеющимися столбцами
    cross_columns_df = df_dbf.parse_columns_df(columns_list=default_columns_list, cross_list=total_columns)

    cross_columns = cross_columns_df['COL_ID'].tolist()
    column_names = cross_columns_df['COL_NAME'].tolist()

    exp1 = exp[cross_columns]

    absbath = os.path.dirname(path)
    basename = os.path.basename(path)
    exportpath = os.path.join(absbath, basename)
    exportpath_csv = f'{exportpath[:-4]}.csv'
    exportpath_xlsx = f'{exportpath[:-4]}.xlsx'
    csv_encoding = "cp1251"

    try:
        to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names, csv_encoding=csv_encoding)
        change_1251_csv_encoding(exportpath_csv)
        os.startfile(exportpath_csv)
        # to_excel_custom_header(df=exp1, excel_path=exportpath_xlsx, column_names=column_names)

    except PermissionError:
        print(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
        logger.error(f"'{basename[:-4]}.csv' is opened, saved as '{basename[:-4]}_1.csv'")
        exportpath = os.path.join(absbath, basename)
        exportpath_csv = f'{exportpath[:-4]}_1.csv'
        exportpath_xlsx = f'{exportpath[:-4]}_1.xlsx'
        to_csv_custom_header(df=exp1, csv_path=exportpath_csv, column_names=column_names, csv_encoding=csv_encoding)
        change_1251_csv_encoding(exportpath_csv)
        os.startfile(exportpath_csv)
        # to_excel_custom_header(df=exp1, excel_path=exportpath_xlsx, column_names=column_names)

    print("~~~ Export Successful ~~~")


# сохраняем в csv c custom столбцами
def to_csv_custom_header(df, csv_path, column_names, csv_encoding):
    print(f'# Info (DF_DBF): Columns saved: {len(column_names)}')
    logger.info(f'# Excel custom header Info (DF_DBF): Columns saved: {len(column_names)}')
    # print(f'# Custom headers info: Columns list: {column_names}')

    with open(csv_path, 'w') as csvfile:
        column_names_string = ''
        for i in column_names:
            column_names_string = column_names_string + f'"{i}",'
        # column_names_string = ",".join([str(i) for i in column_names])
        csvfile.write(f'{column_names_string}\n')

    df.to_csv(csv_path, header=False, index=False, encoding=csv_encoding, mode="a")

def change_1251_csv_encoding(file_path):
    def encode_change(file_path, source_encoding, target_encoding):
        tmp_file = file_path + '.tmp'
        BLOCKSIZE = 1048576  # or some other, desired size in bytes
        with codecs.open(file_path, "r", source_encoding) as sourceFile:
            with codecs.open(tmp_file, "w", target_encoding) as targetFile:
                while True:
                    contents = sourceFile.read(BLOCKSIZE)
                    if not contents:
                        break
                    targetFile.write(contents)
        shutil.move(tmp_file, file_path)

    encode_change(file_path, 'cp1251', 'utf-8')
    encode_change(file_path, 'utf-8', 'utf-8-sig')

# сохраняем в csv c custom столбцами
def to_excel_custom_header(df, excel_path, column_names):
    print(f'# Info (DF_DBF): {len(column_names)}')
    logger.info(f'# Excel cutom header Info (DF_DBF): Columns saved: {len(column_names)}')


    df.columns = column_names
    with pd.ExcelWriter(excel_path) as writer:
        df.to_excel(writer, index=False)


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

    diam = None
    lang = "RU"
    custom_columns = []

    if DEBUG == 1:
        path = r"d:\###WORK\###\123\1nzhu_new_bends.dbf"
    else:
        arg = ''

        if arg != '':
            path = arg
        else:
            path = input("DBF path: ")

        if path[-3:] != 'DBF' and path[-3:] != 'dbf':
            input("##ERROR: not DBF link")
        else:
            export_default(path)

    if DEBUG != 1:
        input("")
    else:
        print("~~~ DEBUG Done ~~~")
