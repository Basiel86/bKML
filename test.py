import simplekml


def simple_kml():
    kml = simplekml.Kml()

    points = [1, 2, 3, 4, 5]
    lats = [56.595274880, 56.595275750, 56.595277900, 56.595278490, 56.595279160]
    longs = [24.171464720, 24.171445160, 24.171394140, 24.171382050, 24.171368290]

    anoms_folder = kml.newfolder(name='Аномалии')
    anoms_points = anoms_folder.newfolder(name='Точки')
    anoms_line = anoms_folder.newfolder(name='Линия')

    ml_folder = kml.newfolder(name='Аномалии')
    ml_points = ml_folder.newfolder(name='Точки')
    ml_line = ml_folder.newfolder(name='Линия')

    i = 0

    line_path = []
    for i in range(len(points)):
        cur_coord = [(longs[i], lats[i])]
        line_path.append((longs[i], lats[i]))
        pnt = anoms_points.newpoint(name=points[i], coords=cur_coord)
        pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
        # pnt = fol.newpoint(name="Точка", coords=[(18.432314, -33.988862)])

    lin = anoms_line.newlinestring(name="point", description="Pipeline", coords=line_path)
    # lin = fol.newlinestring(name="pointЫ", description="Pipeline",
    # coords = [(18.43312, -33.98924), (18.43224, -33.98914),
    #           (18.43144, -33.98911), (18.43095, -33.98904)])
    lin.style.linestyle.color = simplekml.Color.red  # Red
    lin.style.linestyle.width = 3  # 10 pixels

    kml.save("bKML.kml")


if __name__ == '__main__':
    simple_kml()
