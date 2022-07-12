import simplekml
import DF_DBF
import os


ml_range_colors = {'0-10': [99, 190, 123],
                   '10-20': [138, 201, 125],
                   '20-30': [177, 212, 127],
                   '30-40': [216, 223, 129],
                   '40-50': [255, 235, 132],
                   '50-60': [254, 203, 126],
                   '60-70': [252, 170, 120],
                   '70-80': [250, 138, 114],
                   '80-100': [248, 105, 107]}

meter_lang = {'RU': 'м., ',
              'EN': 'm., '}


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
    def __init__(self, dbf_path, diameter, lang="RU"):

        self.dbf_path = dbf_path
        self.lang = lang
        self.diameter = diameter
        self.features_types_list = ['ANOM', 'CONSTRUCT', 'OTHER']
        self.kml = simplekml.Kml()
        self.kml_anomalies_folder = simplekml.Folder
        self.kml_constructions_folder = simplekml.Folder
        self.ml_folder = simplekml.Folder
        self.ml_subranges_folder = simplekml.Folder
        self.kml_others_folder = simplekml.Folder
        self.is_anomalies_folder_exists = False
        self.ml_subranges_folder_exists = False
        self.is_constructions_folder_exists = False
        self.is_others_folder_exists = False

        # self.ml_icon_path = self.kml.addfile(r'icons\ML.png')

    def kml_make_anomalies_folder(self):

        if self.is_anomalies_folder_exists is False:
            self.is_anomalies_folder_exists = True
            if self.lang == "RU":
                fname = "Аномалии"
            else:
                fname = "Anomalies"

            self.kml_anomalies_folder = self.kml.newfolder(name=fname)

    def kml_make_ml_subranges_folder(self):

        if self.ml_subranges_folder_exists is False:
            self.ml_subranges_folder_exists = True
            if self.lang == "RU":
                fname = "Диапазоны глубин"
                ml_fname = "Потеря металла"
            else:
                fname = "Depth ranges"
                ml_fname = "Metal Loss"

            self.ml_folder = self.kml_anomalies_folder.newdocument(name=ml_fname)
            self.ml_subranges_folder = self.ml_folder.newfolder(name=fname)

    def kml_make_constructions_folder(self):

        if self.is_constructions_folder_exists is False:
            self.is_constructions_folder_exists = True
            if self.lang == "RU":
                fname = "Конструктивные элементы"
            else:
                fname = "Construction elements"

            self.kml_constructions_folder = self.kml.newfolder(name=fname)

    def kml_make_others_folder(self):

        if self.is_others_folder_exists is False:
            self.is_others_folder_exists = True

            if self.lang == "RU":
                fname = "Прочие особенности"
            else:
                fname = "Other features"

            self.kml_others_folder = self.kml.newfolder(name=fname)

    def kml_save(self, kml_name):
        # self.kml.save(f"{kml_name}.kml")
        self.kml.savekmz(f"{kml_name}.kmz")

    def kml_write_anomaly(self, feature_name, kml_data, isvisible=None):

        self.kml_make_anomalies_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_anomalies_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
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

    def kml_write_ml_subranges(self, kml_data, isvisible=None):

        max_group_count = 5000

        self.kml_make_anomalies_folder()
        self.kml_make_ml_subranges_folder()

        pname, lname = get_pname_lname(self.lang)

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        feature_line = ([])

        kml_data = kml_data.loc[kml_data['#DEPTH_PRC'] != ""].copy(deep=True)

        for i in range(8):
            min_depth = i * 10
            max_depth = (i + 1) * 10
            ml_range = str(min_depth) + "-" + str(max_depth)

            current_range_df = kml_data[kml_data['#DEPTH_PRC'].between(min_depth, max_depth, inclusive="right")]

            if len(current_range_df) < max_group_count:

                # вытаскивание 5000 глубочайших в серии
                #     current_range_df = current_range_df.sort_values('#DEPTH_PRC')
                #     current_range_df = current_range_df.iloc[0:max_group_count]#.copy(deep=True)
                #     current_range_df = current_range_df.sort_values('#DIST_START')
                # print(f'{min_depth}-{max_depth} count: {len(current_range_df)}')

                if len(current_range_df) > 0:
                    tmp_folder = self.ml_subranges_folder.newfolder(name=f'{i * 10}-{i * 10 + 10}%')
                    tmp_points_folder = tmp_folder.newfolder(name=pname)
                    for elem_id in range(len(current_range_df)):
                        current_coord = [(current_range_df.iloc[elem_id, 2], current_range_df.iloc[elem_id, 1])]
                        feature_line.append((current_range_df.iloc[elem_id, 2], current_range_df.iloc[elem_id, 1]))
                        point = tmp_points_folder.newpoint(name=current_range_df.iloc[elem_id, 0], coords=current_coord,
                                                           visibility=point_visibility)

                        point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes' \
                                                          '/placemark_circle.png '

                        r, g, b = list_to_rgb(ml_range_colors[ml_range])
                        point.style.iconstyle.color = simplekml.Color.rgb(r, g, b)
                        # pnt.description = '<img src="' + path + '" alt="picture" width="400" height="300" align="left" />'

                        # point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_square.png'
                        point.style.labelstyle.scale = 0.9

    def kml_write_construction(self, feature_name, kml_data, isvisible=None):

        self.kml_make_constructions_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_constructions_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
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

    def kml_write_others(self, feature_name, kml_data, isvisible=None):

        self.kml_make_others_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_others_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
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

    def kml_write_top_lvl(self, feature_name, kml_data, isvisible=None):

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
        tmp_points_folder = tmp_folder.newfolder(name=pname)
        # tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(kml_data)):
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1])]
            feature_line.append((kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 1]))
            point = tmp_points_folder.newpoint(name=kml_data.iloc[elem_id, 0], coords=current_coord,
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
            line.style.linestyle.width = 3  # 3 pixels

    def dbf_to_kml(self):

        dbf_conf_init = DF_DBF.df_DBF(DBF_path=self.dbf_path, lang=self.lang)
        df_DBF_raw = dbf_conf_init.convert_dbf(diameter=self.diameter)
        meter = meter_lang[self.lang]


        weld_list = dbf_conf_init.get_welds_list()

        # Оставляем лишь координатные
        df_DBF = df_DBF_raw.loc[df_DBF_raw['#LAT'] != ""].copy(deep=True)

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
                              df_DBF['#HAR_CODE1'].astype(str) + ' ' + \
                              df_DBF['#DESCR'].astype(str)

        df_DBF['WELD_DESCR'] = ''
        df_DBF.loc[df_DBF['#FEA_CODE'].isin(weld_list), 'WELD_DESCR'] = df_DBF['#DIST_START'].astype(str) + meter + \
                                                                       '#' + df_DBF['#JN'].astype(str) + ', ' + \
                                                                       df_DBF['#FEA_CODE_REPLACE'].astype(str) + ' ' + \
                                                                       df_DBF['#HAR_CODE1'].astype(str) + ' ' + \
                                                                       df_DBF['#DESCR'].astype(str)

        anoms_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'ANOM']
        anoms_list = anoms_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in anoms_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            anom_kml_data = current_anom_df[['ML_DESCR', '#LAT', '#LONG', '#DEPTH_PRC', '#DIST_START']]
            if current_anom_name in ['Потеря металла', 'Metal Loss']:
                self.kml_write_ml_subranges(kml_data=anom_kml_data, isvisible=[0, 0])
            else:
                self.kml_write_anomaly(feature_name=current_anom_name, kml_data=anom_kml_data, isvisible=[0, 0])

        other_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'OTH']
        other_list = other_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in other_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            other_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG']]
            self.kml_write_others(feature_name=current_anom_name, kml_data=other_kml_data, isvisible=[0, 0])

        construct_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'CON']
        construct_list = construct_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in construct_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]
            construct_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG']]

            if current_anom_name in ['Задвижка', 'Valve']:
                isvisible = [1, 0]
            else:
                isvisible = [0, 0]
            self.kml_write_construction(feature_name=current_anom_name, kml_data=construct_kml_data,
                                        isvisible=isvisible)

        top_df = df_DBF.loc[df_DBF['#KML_CLASS'] == 'TOP']
        top_list = top_df['#FEA_CODE_REPLACE'].value_counts(ascending=True)
        for current_anom_name, row in top_list.items():
            current_anom_df = df_DBF.loc[df_DBF['#FEA_CODE_REPLACE'] == current_anom_name]

            if current_anom_name in ['Шов', 'Weld']:
                top_kml_data = current_anom_df[["WELD_DESCR", '#LAT', '#LONG']]
                isvisible = [0, 1]
            else:
                top_kml_data = current_anom_df[["OTH_DESCR", '#LAT', '#LONG']]
                isvisible = [1, 0]

            self.kml_write_top_lvl(feature_name=current_anom_name, kml_data=top_kml_data, isvisible=isvisible)

        absbath = os.path.dirname(self.dbf_path)
        basename = os.path.basename(self.dbf_path)
        exportpath = os.path.join(absbath, basename)
        exportpath = f'{exportpath[:-4]}'

        self.kml_save(exportpath)

        no_coord_records = len(df_DBF_raw) - len(df_DBF)
        print("\n~~~Done~~~")
        print(f"Обработано записей: {len(df_DBF_raw)}")
        if no_coord_records > 0:
            print("Записей без координат: ", no_coord_records)
        # print(f"KML создана по {len(df_DBF)} записям"'\n')


def main():
    lang = 'RU'
    path = input("Enter DBF path: ")
    diameter = float(input("Enter Diameter: "))

    # path = r'd:\WORK\#Thailand\NXB 14 inch MTP - SRS, 62 km\Reports\FR\Database\1nxbu.DBF'
    # path = r'd:\WORK\#Thailand\NXD 18 inch  LLK to SRB, 93 km\Reports\FR\Database\Parts_OK\1nxdu_1.dbf'
    # diameter = 11
    # bKML(dbf_path=path, lang=lang, diameter=diameter).dbf_to_kml()

    try:
        bKML(dbf_path=path, lang=lang, diameter=diameter).dbf_to_kml()
        input("~~~Done~~~")
    except Exception as ex:
        print(ex)
        input("Что-то пошло не так...")


if __name__ == "__main__":
    main()
