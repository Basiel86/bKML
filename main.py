import simplekml
import DF_DBF


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
    def __init__(self, lang):

        self.lang = lang
        self.features_types_list = ['ANOM', 'CONSTRUCT', 'OTHER']
        self.kml = simplekml.Kml()
        self.kml_anomalies_folder = simplekml.Folder
        self.kml_constructions_folder = simplekml.Folder
        self.kml_others_folder = simplekml.Folder
        self.is_anomalies_folder_exists = False
        self.is_constructions_folder_exists = False
        self.is_others_folder_exists = False

    def kml_make_anomalies_folder(self):

        if self.is_anomalies_folder_exists is False:
            self.is_anomalies_folder_exists = True
            if self.lang == "RU":
                fname = "Аномалии"
                pname, lname = get_pname_lname(self.lang)
            else:
                fname = "Anomalies"
                pname, lname = get_pname_lname(self.lang)

            self.kml_anomalies_folder = self.kml.newfolder(name=fname)

    def kml_make_constructions_folder(self):

        if self.is_constructions_folder_exists is False:
            self.is_constructions_folder_exists = True
            if self.lang == "RU":
                fname = "Конструктивные элементы"
                pname, lname = get_pname_lname()
            else:
                fname = "Construction elements"
                pname, lname = get_pname_lname()

            self.kml_constructions_folder = self.kml.newfolder(name=fname)

    def kml_make_others_folder(self):

        if self.is_others_folder_exists is False:
            self.is_others_folder_exists = True

            if self.lang == "RU":
                fname = "Прочие особенности"
                pname, lname = get_pname_lname(self.lang)
            else:
                fname = "Other features"
                pname, lname = get_pname_lname(self.lang)

            self.kml_others_folder = self.kml.newfolder(name=fname)

    def kml_save(self, kml_name):
        self.kml.save(f"{kml_name}.kml")

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


def main():
    toKML = bKML(lang="EN")

    path = input("Enter DBF path: ")

    # path = r"1nopm.DBF"

    dbf_conv = DF_DBF.df_DBF(DBF_path=path, lang="EN")
    df_DBF = dbf_conv.convert_dbf(diameter=10.75)

    # формируем описание для потерей
    df_DBF['ML_DESCR'] = ''
    df_DBF.loc[df_DBF['FEA_DEPTH_PRC'] != "", 'ML_DESCR'] = df_DBF['FEA_DIST'].astype(str) + 'м., ' + df_DBF[
        'Feature'] + ' ' + df_DBF['FEA_DEPTH_PRC'].astype(str) + '%'
    df_DBF.loc[df_DBF['FEA_DEPTH_PRC'] == "", 'ML_DESCR'] = df_DBF['FEA_DIST'].astype(str) + 'м., ' + df_DBF['Feature']

    # описание для всего остального
    df_DBF['OTH_DESCR'] = df_DBF['FEA_DIST'].astype(str) + 'м., ' + df_DBF['Feature'] + ' ' + df_DBF['COMMENT']

    anoms_df = df_DBF.loc[df_DBF['CLASS'] == 'ANOM']
    anoms_list = anoms_df['Feature'].value_counts(ascending=True)
    for current_anom_name, row in anoms_list.items():
        current_anom_df = df_DBF.loc[df_DBF['Feature'] == current_anom_name]
        anom_kml_data = current_anom_df[["ML_DESCR", 'LATITUDE', 'LONGITUDE']]
        toKML.kml_write_anomaly(feature_name=current_anom_name, kml_data=anom_kml_data, isvisible=[0, 0])

    other_df = df_DBF.loc[df_DBF['CLASS'] == 'OTH']
    other_list = other_df['Feature'].value_counts(ascending=True)
    for current_anom_name, row in other_list.items():
        current_anom_df = df_DBF.loc[df_DBF['Feature'] == current_anom_name]
        other_kml_data = current_anom_df[["OTH_DESCR", 'LATITUDE', 'LONGITUDE']]
        toKML.kml_write_others(feature_name=current_anom_name, kml_data=other_kml_data, isvisible=[0, 0])

    construct_df = df_DBF.loc[df_DBF['CLASS'] == 'CON']
    construct_list = construct_df['Feature'].value_counts(ascending=True)
    for current_anom_name, row in construct_list.items():
        current_anom_df = df_DBF.loc[df_DBF['Feature'] == current_anom_name]
        construct_kml_data = current_anom_df[["OTH_DESCR", 'LATITUDE', 'LONGITUDE']]

        if current_anom_name in ['Задвижка', 'Valve']:
            isvisible = [1, 0]
        else:
            isvisible = [0, 0]
        toKML.kml_write_construction(feature_name=current_anom_name, kml_data=construct_kml_data, isvisible=isvisible)

    top_df = df_DBF.loc[df_DBF['CLASS'] == 'TOP']
    top_list = top_df['Feature'].value_counts(ascending=True)
    for current_anom_name, row in top_list.items():
        current_anom_df = df_DBF.loc[df_DBF['Feature'] == current_anom_name]
        top_kml_data = current_anom_df[["OTH_DESCR", 'LATITUDE', 'LONGITUDE']]

        if current_anom_name in ['Шов', 'Weld']:
            isvisible = [0, 1]
        else:
            isvisible = [1, 0]
        toKML.kml_write_top_lvl(feature_name=current_anom_name, kml_data=top_kml_data, isvisible=isvisible)

    toKML.kml_save("Тест")


if __name__ == "__main__":
    main()
