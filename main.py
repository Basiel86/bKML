import sys
import dbf
import simplekml

kml = simplekml.Kml()

# SOURCE_NAME = input("Enter full DBF path: ")

kml_welds_folder = kml.newfolder(name='Швы')
kml_welds_points_folder = kml_welds_folder.newfolder(name='Точки')
kml_welds_line_folder = kml_welds_folder.newfolder(name='Линия')

kml_anomalies_folder = kml.newfolder(name='Аномалии')
kml_anomalies_points_folder = kml_anomalies_folder.newfolder(name='Точки')
kml_anomalies_line_folder = kml_anomalies_folder.newfolder(name='Линия')

kml_ml_folder_folder = kml.newfolder(name='Потери металла')
kml_ml_points_folder = kml_ml_folder_folder.newfolder(name='Точки')
kml_ml_line_folder = kml_ml_folder_folder.newfolder(name='Линия')

kml_geom_folder = kml.newfolder(name='Дефекты геометрии')
kml_geom_points_folder = kml_geom_folder.newfolder(name='Точки')
kml_geom_line_folder = kml_geom_folder.newfolder(name='Линия')

WELDS_ID = [901, 902, 903, 904, 905, 906]
ML_ID = [101, 106, 32869, 3581198, 1081016421, 111]
GEOM_ID = [201, 202]

welds_point_names = []
welds_line = []
welds_lats = []
welds_longs = []


ml__point_names = []
ml_line = []
ml_lats = []
ml_longs = []

geom_point_names = []
geom_lines = []
geom_lats = []
geom_longs = []

SOURCE_NAME = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\1nxeu.dbf"

extention = SOURCE_NAME[-3:]

if extention != 'dbf' and extention != 'DBF':
    print("Wrong file Path")
    input("Click any key for exit ")
    sys.exit()

newFormatRECOPY_SOURCE = dbf.Table(SOURCE_NAME, codepage='cp1251', on_disk=True)
newFormatRECOPY_SOURCE.open(dbf.READ_WRITE)

with newFormatRECOPY_SOURCE:

    for record in newFormatRECOPY_SOURCE:
        if record.FEA_CODE in WELDS_ID:
            welds_point_names.append(record.FEA_DIST)
            welds_longs.append(record.LONGITUDE)
            welds_lats.append(record.LATITUDE)
        if record.FEA_CODE in ML_ID:
            ml__point_names.append(record.FEA_DIST)
            ml_longs.append(record.LONGITUDE)
            ml_lats.append(record.LATITUDE)
        if record.FEA_CODE in GEOM_ID:
            geom_point_names.append(record.FEA_DIST)
            geom_longs.append(record.LONGITUDE)
            geom_lats.append(record.LATITUDE)

for i in range(len(welds_point_names)):
    cur_coord = [(welds_longs[i], welds_lats[i])]
    welds_line.append((welds_longs[i], welds_lats[i]))
    pnt = kml_welds_points_folder.newpoint(name=welds_point_names[i], coords=cur_coord)
    pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'


lin = kml_welds_line_folder.newlinestring(name="Профиль трубных секций", description="Трубные секции", coords=welds_line)
lin.style.linestyle.color = simplekml.Color.red  # Red
lin.style.linestyle.width = 3  # 10 pixels

kml.save("bKML.kml")
