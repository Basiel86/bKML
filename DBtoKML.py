import simplekml
import DF_DBF
import os
import traceback
from datetime import datetime

import logging

import parse_inch

logger = logging.getLogger('app.DBtoKML')

ml_range_colors = {'0-10': [99, 190, 123],
                   '10-20': [138, 201, 125],
                   '20-30': [177, 212, 127],
                   '30-40': [216, 223, 129],
                   '40-50': [255, 235, 132],
                   '50-60': [254, 203, 126],
                   '60-70': [252, 170, 120],
                   '70-80': [250, 138, 114],
                   '80-90': [248, 105, 107],
                   '90-100': [248, 105, 107]}

meter_lang = {'RU': 'м., ',
              'EN': 'm., '}

int_list = ['Internal', 'Внутренний', 'Interno']
ext_list = ['External', 'Внешний', 'Externo']


def list_to_rgb(rgb_list):
    return rgb_list[0], rgb_list[1], rgb_list[2]


def get_pname_lname(lang="EN"):
    if lang == "RU":
        pname = "Точки"
        lname = "Профиль"
        return pname, lname
    else:
        pname = "Points"
        lname = "Line"
        return pname, lname


class bKML:
    def __init__(self):

        self.lang = 'RU'
        self.kml = simplekml.Kml()
        self.kml_anomalies_folder = simplekml.Folder
        self.kml_constructions_folder = simplekml.Folder
        self.ml_folder = simplekml.Folder
        self.ml_int_folder = simplekml.Folder
        self.ml_ext_folder = simplekml.Folder
        self.ml_subranges_folder = simplekml.Folder
        self.kml_others_folder = simplekml.Folder
        self.is_anomalies_folder_exists = False
        self.is_ml_folder_exist = False
        self.is_constructions_folder_exists = False
        self.is_others_folder_exists = False
        self.is_int_ml_folder_exists = False
        self.is_ext_ml_folder_exists = False
        self.line_width = 3

        # self.ml_icon_path = self.kml.addfile(r'icons\ML.png')

    def __kml_make_anomalies_folder(self):

        if self.is_anomalies_folder_exists is False:
            self.is_anomalies_folder_exists = True
            if self.lang == "RU":
                fname = "Аномалии"
            else:
                fname = "Anomalies"

            self.kml_anomalies_folder = self.kml.newfolder(name=fname)

    def __kml_make_ml_folder(self, ml_type=None):

        if self.is_ml_folder_exist is False:
            self.is_ml_folder_exist = True
            if self.lang == "RU":
                ml_fname = "Потеря металла"
            else:
                ml_fname = "Metal Loss"

            self.ml_folder = self.kml_anomalies_folder.newdocument(name=ml_fname)

        if ml_type in int_list:
            if self.is_int_ml_folder_exists is False:
                self.ml_int_folder = self.ml_folder.newfolder(name=ml_type)
        if ml_type in ext_list:
            if self.is_ext_ml_folder_exists is False:
                self.ml_ext_folder = self.ml_folder.newfolder(name=ml_type)

    def __kml_make_constructions_folder(self):

        if self.is_constructions_folder_exists is False:
            self.is_constructions_folder_exists = True
            if self.lang == "RU":
                fname = "Конструктивные элементы"
            else:
                fname = "Construction elements"

            self.kml_constructions_folder = self.kml.newfolder(name=fname)

    def __kml_make_others_folder(self):

        if self.is_others_folder_exists is False:
            self.is_others_folder_exists = True

            if self.lang == "RU":
                fname = "Прочие особенности"
            else:
                fname = "Other features"

            self.kml_others_folder = self.kml.newfolder(name=fname)

    def __kml_save(self, kml_name):
        # self.kml.save(f"{kml_name}.kml")
        self.kml.savekmz(f"{kml_name}.kmz")

    def __kml_write_anomaly(self, feature_name, kml_data, isvisible=None, add_points_folder=False):

        self.__kml_make_anomalies_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_anomalies_folder.newdocument(name=feature_name)
        if add_points_folder is True:
            tmp_points_folder = tmp_folder.newfolder(name=pname)
        else:
            tmp_points_folder = tmp_folder
        # tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(kml_data)):
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1])]
            feature_line.append((kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1]))
            point = tmp_points_folder.newpoint(name=kml_data.iloc[elem_id, 0], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_square.png'
            point.style.labelstyle.scale = 0.9

        # line = tmp_lines_folder.newlinestring(name="line", description="line",
        #                                       coords=feature_line, visibility=line_visibility)
        # line.style.linestyle.color = simplekml.Color.red  # Red
        # line.style.linestyle.width = 3  # 3 pixels

    def __kml_write_ml_subranges(self, kml_data, isvisible=None, add_points_folder=False):

        max_group_count = 5000

        self.__kml_make_anomalies_folder()

        pname, lname = get_pname_lname(self.lang)

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        feature_line = ([])

        kml_data = kml_data.loc[kml_data['#DEPTH_PRC'] != ""].copy(deep=True)

        # список доступных типов потерей
        loc_vals = kml_data['#LOC'].value_counts(dropna=False, normalize=False)

        # бежим по типам
        for index, value in loc_vals.items():
            # если тип в списке внутренних или внешних то входим
            if index in int_list or index in ext_list:

                # текущие Тип и Датафрэйм
                current_ml_type = index
                current_ml_type_df = kml_data.loc[kml_data['#LOC'] == current_ml_type].copy(deep=True)

                # создаем папку с ПМ и типом
                self.__kml_make_ml_folder(ml_type=current_ml_type)

                # создаем временную папку отвечающую за тип потери
                if current_ml_type in int_list:
                    tmp_current_type_ml_folder = self.ml_int_folder
                else:
                    tmp_current_type_ml_folder = self.ml_ext_folder

                # пишем диапазоны
                for i in range(10):
                    min_depth = i * 10
                    max_depth = (i + 1) * 10
                    ml_range = str(min_depth) + "-" + str(max_depth)

                    current_range_df = current_ml_type_df[
                        current_ml_type_df['#DEPTH_PRC'].between(min_depth, max_depth, inclusive="left")]

                    if len(current_range_df) < max_group_count:

                        # вытаскивание 5000 глубочайших в серии
                        #     current_range_df = current_range_df.sort_values('#DEPTH_PRC')
                        #     current_range_df = current_range_df.iloc[0:max_group_count]#.copy(deep=True)
                        #     current_range_df = current_range_df.sort_values('#DIST_START')
                        # print(f'{min_depth}-{max_depth} count: {len(current_range_df)}')

                        if len(current_range_df) > 0:
                            tmp_folder = tmp_current_type_ml_folder.newfolder(name=f'{i * 10}-{i * 10 + 10}%')

                            if add_points_folder is True:
                                tmp_points_folder = tmp_folder.newfolder(name=pname)
                            else:
                                tmp_points_folder = tmp_folder

                            for elem_id in range(len(current_range_df)):
                                current_coord = [(current_range_df.iloc[elem_id, 2], current_range_df.iloc[elem_id, 1])]
                                feature_line.append(
                                    (current_range_df.iloc[elem_id, 2], current_range_df.iloc[elem_id, 1]))
                                point = tmp_points_folder.newpoint(name=current_range_df.iloc[elem_id, 0],
                                                                   coords=current_coord,
                                                                   visibility=point_visibility)

                                point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes' \
                                                                  '/placemark_circle.png '

                                r, g, b = list_to_rgb(ml_range_colors[ml_range])
                                point.style.iconstyle.color = simplekml.Color.rgb(r, g, b)
                                # pnt.description = '<img src="' + path + '" alt="picture" width="400" height="300" align="left" />'
                                # point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_square.png'
                                point.style.labelstyle.scale = 0.9

    def __kml_write_construction(self, feature_name, kml_data, isvisible=None, add_points_folder=False):

        self.__kml_make_constructions_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_constructions_folder.newdocument(name=feature_name)

        if add_points_folder is True:
            tmp_points_folder = tmp_folder.newfolder(name=pname)
        else:
            tmp_points_folder = tmp_folder

        # tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(kml_data)):
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1])]
            feature_line.append((kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1]))
            point = tmp_points_folder.newpoint(name=kml_data.iloc[elem_id, 0], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
            point.style.labelstyle.scale = 0.9

        # line = tmp_lines_folder.newlinestring(name="line", description="line",
        #                                       coords=feature_line, visibility=line_visibility)
        # line.style.linestyle.color = simplekml.Color.red  # Red
        # line.style.linestyle.width = 3  # 3 pixels

    def __kml_write_others(self, feature_name, kml_data, isvisible=None, add_points_folder=False):

        self.__kml_make_others_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_others_folder.newdocument(name=feature_name)

        if add_points_folder is True:
            tmp_points_folder = tmp_folder.newfolder(name=pname)
        else:
            tmp_points_folder = tmp_folder
        # tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(kml_data)):
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1])]
            feature_line.append((kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1]))
            point = tmp_points_folder.newpoint(name=kml_data.iloc[elem_id, 0], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
            point.style.labelstyle.scale = 0.9

        # line = tmp_lines_folder.newlinestring(name="line", description="line",
        #                                       coords=feature_line, visibility=line_visibility)
        # line.style.linestyle.color = simplekml.Color.red  # Red
        # line.style.linestyle.width = 3  # 3 pixels

    def __kml_write_top_lvl(self, feature_name, kml_data, isvisible=None, points_folder=False):

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        if self.lang == "RU":
            if feature_name in ['Шов', 'Weld']:
                feature_name = "Трубные секции"
        else:
            if feature_name in ['Шов', 'Weld']:
                feature_name = "Weld sections"

        tmp_folder = self.kml.newdocument(name=feature_name)
        tmp_folder.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

        if points_folder is True:
            tmp_points_folder = tmp_folder.newfolder(name=pname)
        else:
            tmp_points_folder = tmp_folder
        # tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(kml_data)):

            point_descr = kml_data.iloc[elem_id, 0]
            point_remarks = kml_data.iloc[elem_id, 3]

            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1])]

            # если коммента - Отвод Середина нет ни в десрипшне ни в ремарке - исключаем из использования в линии
            # отвод середину скипаем из линии
            if 'отвод середина' not in str(point_descr).lower() and 'отвод середина' not in str(point_remarks).lower():
                feature_line.append((kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1]))

            # точки с GPS TMP не пишем
            if 'отвод' not in str(point_descr).lower() and 'отвод' not in str(point_remarks).lower():
                point = tmp_points_folder.newpoint(name=point_descr, coords=current_coord,
                                                   visibility=point_visibility)
                if feature_name in ['Маркер', 'Marker']:
                    point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/flag.png'
                else:
                    point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
                point.style.labelstyle.scale = 0.9
                point.style.iconstyle.scale = 0.8

        if line_visibility == 1:
            tmp_lines_folder = tmp_folder.newfolder(name=lname)
            line = tmp_lines_folder.newlinestring(name="line", description="line",
                                                  coords=feature_line, visibility=line_visibility)
            line.style.linestyle.color = simplekml.Color.rgb(217, 0, 0)
            line.style.linestyle.width = self.line_width  # 3 pixels

    @staticmethod
    def dbf_load(dbf_path, diameter, lang="RU") -> tuple:

        """
        Конвертим DBF в DF
        :return: список (Класс DF_DBF, DF, dbf_path)
        """

        dbf_conf_init = DF_DBF.df_DBF(local=True)
        cls = dbf_conf_init
        df = dbf_conf_init.convert_dbf(diameter=diameter, dbf_path=dbf_path, lang=lang)
        return cls, df, dbf_path

    def dbf_to_kml_fast(self, dbf_path, line_width):

        logger.info("DBF to KML Fast")
        print(f'# Info (KML): DBF to KML Fast: {dbf_path}')


        diameter = parse_inch.parse_inch_prj(dbf_path)
        if diameter is None:
            diameter = float(input("Enter Diameter: "))

        cls, df, dbf_path = self.dbf_load(dbf_path=dbf_path, lang='RU', diameter=diameter)
        kml_path = self.dbf_to_kml(line_width, df=df, df_dbf_class=cls, export_path=dbf_path)
        return kml_path

    def dbf_to_kml(self, line_width, df_dbf_class, df, export_path):

        """
        конвертируем Датафрейм в KML
        :param line_width:
        :param df_dbf_class:
        :param df:
        :param export_path:
        """

        self.line_width = line_width

        df_DBF_raw = df
        dbf_conf_init = df_dbf_class

        # не берем в расчет отремонтированные
        # df_DBF_raw = df_DBF_raw[df_DBF_raw['#DOC'] != 'T']

        meter = meter_lang[self.lang]

        weld_list = dbf_conf_init.get_welds_list()

        # Оставляем лишь координатные
        df_DBF = df_DBF_raw.loc[df_DBF_raw['#LAT'] != ""].copy(deep=True)

        # записи без координат, если пусто - выходим
        num_of_no_coord_records = len(df_DBF_raw) - len(df_DBF)
        if len(df_DBF) == 0:
            print("# Info (KML): No features with coordinates found")
            logger.warning('No features with coordinates found')
            return None

        print(f"# Info (KML): Total records with NO coordinates: {num_of_no_coord_records}")
        logger.info(f"Total records with NO coordinates: {num_of_no_coord_records}")

        # print (df_DBF['#FEA_NUM'])

        # формируем описание для аномалий
        df_DBF['ML_DESCR'] = ''
        df_DBF.loc[df_DBF['#DEPTH_PRC'] != "", 'ML_DESCR'] = df_DBF['#DIST_START'].astype(str) + meter + \
                                                             '№' + df_DBF['#FEA_NUM'].astype(str) + ' ' + \
                                                             df_DBF['#FEA_CODE_REPLACE'].astype(str) + ' ' + \
                                                             df_DBF['#DEPTH_PRC'].astype(str) + '%'
        df_DBF.loc[df_DBF['#DEPTH_PRC'] == "", 'ML_DESCR'] = df_DBF['#DIST_START'].astype(str) + meter + \
                                                             '№' + df_DBF['#FEA_NUM'].astype(str) + ' ' + \
                                                             df_DBF['#FEA_CODE_REPLACE'].astype(str)

        # описание для всего остального
        df_DBF['OTH_DESCR'] = df_DBF['#DIST_START'].astype(str) + meter + \
                              '№' + df_DBF['#FEA_NUM'].astype(str) + ' ' + \
                              df_DBF['#FEA_CODE_REPLACE'].astype(str) + ' ' + \
                              df_DBF['#HAR_CODE1_REPLACE'].astype(str) + ' ' + \
                              df_DBF['#DBF_DESCR'].astype(str)

        df_DBF['WELD_DESCR'] = ''
        df_DBF.loc[df_DBF['#FEA_CODE'].isin(weld_list), 'WELD_DESCR'] = df_DBF['#DIST_START'].astype(str) + meter + \
                                                                        '#' + df_DBF['#JN'].astype(str) + ', ' + \
                                                                        df_DBF['#FEA_CODE_REPLACE'].astype(str) + ' ' + \
                                                                        df_DBF['#HAR_CODE1_REPLACE'].astype(str) + ' ' + \
                                                                        df_DBF['#DBF_DESCR'].astype(str)

        anoms_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'ANOM']
        anoms_list = anoms_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in anoms_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            anom_kml_data = current_anom_df[
                ['ML_DESCR', '#LAT', '#LONG', '#DEPTH_PRC', '#DIST_START', '#LOC', '#CORR', '#RWT', '#ERF']]
            if current_anom_name in ['Потеря металла', 'Metal Loss', 'Pérdida de metal']:
                self.__kml_write_ml_subranges(kml_data=anom_kml_data, isvisible=[0, 0])
            else:
                self.__kml_write_anomaly(feature_name=current_anom_name, kml_data=anom_kml_data, isvisible=[0, 0])

        other_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'OTH']
        other_list = other_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in other_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            other_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG']]
            self.__kml_write_others(feature_name=current_anom_name, kml_data=other_kml_data, isvisible=[0, 0])

        construct_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'CON']
        construct_list = construct_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in construct_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            construct_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG']]

            if current_anom_name in ['Задвижка (шток)', 'Valve']:
                isvisible = [1, 0]
            else:
                isvisible = [0, 0]
            self.__kml_write_construction(feature_name=current_anom_name, kml_data=construct_kml_data,
                                          isvisible=isvisible)

        top_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'TOP']
        top_list = top_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in top_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]

            if current_anom_name in ['Шов', 'Weld']:
                top_kml_data = current_anom_df[["WELD_DESCR", '#LAT', '#LONG', '#REMARKS']]
                isvisible = [0, 1]
                points_folder = True
            else:
                top_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG', '#REMARKS']]
                isvisible = [1, 0]
                points_folder = False

            # add_to_line = True if current_anom_name == "GPS TMP" else False

            self.__kml_write_top_lvl(feature_name=current_anom_name, kml_data=top_kml_data,
                                     isvisible=isvisible, points_folder=points_folder)

        absbath = os.path.dirname(export_path) + '\KML'

        if not os.path.exists(absbath):
            os.mkdir(absbath)

        basename = os.path.basename(export_path)
        exportpath = os.path.join(absbath, basename)
        exportpath = f'{exportpath[:-4]}'

        now = datetime.now()
        dt_string = now.strftime("%Y.%m.%d %H.%M.%S")
        exportpath = f'{exportpath} - {dt_string} - {len(df_DBF_raw)} records'

        self.__kml_save(exportpath)

        print(f"# Info (KML): Обработано записей: {len(df_DBF_raw)}")
        logger.info(f"Обработано записей: {len(df_DBF_raw)}")

        return exportpath + '.kmz'

        # print(f"KML создана по {len(df_DBF)} записям"'\n')


def __main():
    DEBUG = 1

    kml_class = bKML()

    if DEBUG == 1:
        path = r'd:\WORK\Beloyarskneftegaz\NBF 12 inch НП п.Андра-НПС Красноленинская (подводный переход) основная нитка, 15 km\Report\FR\DataBase\111\1nbfm.DBF'
        # path = r'd:\###WORK\###\123\1nzhu_new_bends.dbf'

        kml_class.dbf_to_kml_fast(path, 3)

        # diameter = 11
        # lang = 'RU'
        # cls, df, dbf_path = kml_class.dbf_load(dbf_path=path, lang=lang, diameter=diameter)
        # kml_class.dbf_to_kml(line_width=3, df=df, df_dbf_class=cls, export_path=dbf_path)
    else:
        try:
            lang = 'RU'
            path = input("Enter DBF path: ")
            diameter = float(input("Enter Diameter: "))

            cls, df, dbf_path = kml_class.dbf_load(dbf_path=path, lang=lang, diameter=diameter)
            kml_class.dbf_to_kml(line_width=3, df=df, df_dbf_class=cls, export_path=dbf_path)
            input("~~~Done~~~\n")
        except Exception as ex:
            print(ex)
            print(traceback.format_exc())
            input("Что-то пошло не так...")


if __name__ == "__main__":
    __main()
