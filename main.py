import os.path
import pathlib
import sys
import dbf
import simplekml
import pandas as pd
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

    def kml_make_constructions_folder(self, lang="RU"):

        if self.is_constructions_folder_exists is False:
            self.is_constructions_folder_exists = True
            if lang == "RU":
                fname = "Конструктивные элементы"
                pname, lname = get_pname_lname(lang)
            else:
                fname = "Construction elements"
                pname, lname = get_pname_lname(lang)

            self.kml_constructions_folder = self.kml.newfolder(name=fname)

    def kml_make_others_folder(self, lang="RU"):

        if self.is_others_folder_exists is False:
            self.is_others_folder_exists = True

            if lang == "RU":
                fname = "Прочие особенности"
                pname, lname = get_pname_lname(lang)
            else:
                fname = "Other features"
                pname, lname = get_pname_lname(lang)

            self.kml_others_folder = self.kml.newfolder(name=fname)

    def kml_save(self, kml_name):
        self.kml.save(f"{kml_name}.kml")

    def kml_write_anomaly(self, feature_name, kml_data, isvisible=None):

        description_tuple = kml_data[0]
        latitude_tuple = kml_data[1]
        longitude_tuple = kml_data[2]

        self.kml_make_anomalies_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_anomalies_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
        tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id, row in kml_data.iterrows():
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 2])]
            feature_line.append((longitude_tuple[elem_id], latitude_tuple[elem_id]))
            point = tmp_points_folder.newpoint(name=description_tuple[elem_id], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

        line = tmp_lines_folder.newlinestring(name="line", description="line",
                                              coords=feature_line, visibility=line_visibility)
        line.style.linestyle.color = simplekml.Color.red  # Red
        line.style.linestyle.width = 3  # 3 pixels

    def kml_write_construction(self, feature_name, kml_data, isvisible=None):

        description_tuple = kml_data[0]
        latitude_tuple = kml_data[1]
        longitude_tuple = kml_data[2]

        self.kml_make_constructions_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_constructions_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
        tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(description_tuple)):
            current_coord = [(longitude_tuple[elem_id], latitude_tuple[elem_id])]
            feature_line.append((longitude_tuple[elem_id], latitude_tuple[elem_id]))
            point = tmp_points_folder.newpoint(name=description_tuple[elem_id], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

        line = tmp_lines_folder.newlinestring(name="line", description="line",
                                              coords=feature_line, visibility=line_visibility)
        line.style.linestyle.color = simplekml.Color.red  # Red
        line.style.linestyle.width = 3  # 3 pixels

    def kml_write_others(self, feature_name, kml_data, isvisible=None):

        description_tuple = kml_data[0]
        latitude_tuple = kml_data[1]
        longitude_tuple = kml_data[2]

        self.kml_make_others_folder()

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml_others_folder.newdocument(name=feature_name)
        tmp_points_folder = tmp_folder.newfolder(name=pname)
        tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id in range(len(description_tuple)):
            current_coord = [(longitude_tuple[elem_id], latitude_tuple[elem_id])]
            feature_line.append((longitude_tuple[elem_id], latitude_tuple[elem_id]))
            point = tmp_points_folder.newpoint(name=description_tuple[elem_id], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

        line = tmp_lines_folder.newlinestring(name="line", description="line",
                                              coords=feature_line, visibility=line_visibility)
        line.style.linestyle.color = simplekml.Color.red  # Red
        line.style.linestyle.width = 3  # 3 pixels

    def kml_write_top_lvl(self, feature_name, kml_data, isvisible=None):

        description_tuple = kml_data[0]
        latitude_tuple = kml_data[1]
        longitude_tuple = kml_data[2]

        if isvisible is None:
            isvisible = [1, 0]

        point_visibility = isvisible[0]
        line_visibility = isvisible[1]
        feature_line = ([])

        pname, lname = get_pname_lname(self.lang)

        tmp_folder = self.kml.newdocument(name=feature_name)
        tmp_folder.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
        tmp_points_folder = tmp_folder.newfolder(name=pname)
        tmp_lines_folder = tmp_folder.newfolder(name=lname)

        for elem_id, row in kml_data.iterrows():
            current_coord = [(kml_data.iloc[elem_id, 2], kml_data.iloc[elem_id, 2])]
            feature_line.append((longitude_tuple[elem_id], latitude_tuple[elem_id]))
            point = tmp_points_folder.newpoint(name=description_tuple[elem_id], coords=current_coord,
                                               visibility=point_visibility)
            point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

        line = tmp_lines_folder.newlinestring(name="line", description="line",
                                              coords=feature_line, visibility=line_visibility)
        line.style.linestyle.color = simplekml.Color.red  # Red
        line.style.linestyle.width = 3  # 3 pixels


def main():
    path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\1nldm.dbf"

    dbf_conv = DF_DBF.df_DBF(DBF_path=path)
    df_DBF = dbf_conv.convert_dbf()
    kml_data = df_DBF[["Feature", 'LATITUDE', 'LONGITUDE']]

    bKML.kml_write_anomaly(feature_name="Потери металла", kml_data=kml_data)


if __name__ == "__main__":
    main()

# WELDS_ID = [901, 902, 903, 904, 905, 906]
# ML_ID = [101, 106, 32869, 3581198, 1081016421, 111]
# GEOM_ID = [201, 202]
#
# welds_point_names = []
# welds_line = []
# welds_lats = []
# welds_longs = []
#
# ml_point_names = []
# ml_line = []
# ml_lats = []
# ml_longs = []
#
# geom_point_names = []
# geom_lines = []
# geom_lats = []
# geom_longs = []
#
# SOURCE_NAME = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\1ntom.DBF"
#
# extention = SOURCE_NAME[-3:]
#
# if extention != 'dbf' and extention != 'DBF':
#     print("Wrong file Path")
#     input("Click any key for exit ")
#     sys.exit()
#
# newFormatRECOPY_SOURCE = dbf.Table(SOURCE_NAME, codepage='cp1251', on_disk=True)
# newFormatRECOPY_SOURCE.open(dbf.READ_WRITE)
#
# lang = "EN"
#
# bKML = bKML(lang=lang)
#
# with newFormatRECOPY_SOURCE:
#     for record in newFormatRECOPY_SOURCE:
#
#         lat = record.LATITUDE
#         long = record.LONGITUDE
#         dist = record.FEA_DIST
#         depth = record.FEA_DEPTH
#         wt = record.WT
#
#         if depth is not None and wt is not None:
#             depth = round(record.FEA_DEPTH / record.WT * 100, 1)
#         ml_description = f"Потеря металла, {dist}м, D {depth}%"
#         geom_description = f"Вмятина, {dist}м, D {depth}%"
#
#         if record.FEA_CODE in WELDS_ID:
#             welds_point_names.append(dist)
#             welds_longs.append(long)
#             welds_lats.append(lat)
#         if record.FEA_CODE in ML_ID:
#             ml_point_names.append(ml_description)
#             ml_longs.append(long)
#             ml_lats.append(lat)
#         if record.FEA_CODE in GEOM_ID:
#             geom_point_names.append(dist)
#             geom_longs.append(long)
#             geom_lats.append(lat)
#
# bKML.kml_write_anomaly(feature_name="Потери металла", kml_data=[ml_point_names, ml_lats, ml_longs],
#                        isvisible=[1, 0])
# bKML.kml_write_anomaly(feature_name="Вмятины",
#                        kml_data=[geom_point_names, geom_lats, geom_longs], isvisible=[1, 0])
# bKML.kml_write_top_lvl(feature_name="Трубные секции",
#                        kml_data=[welds_point_names, welds_lats, welds_longs], isvisible=[0, 1])
#
# bKML.kml_save("bKML_test")
#
# print("~~~Done~~~")
