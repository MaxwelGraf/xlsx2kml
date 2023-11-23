import pandas as pd
import simplekml
from geopy.distance import geodesic  # расчет расстояния
#from pathlib import Path

#****относительный путь
import os
file_path_credentials = os.path.join(os.path.dirname(__file__), 'data_input.xlsx')
with open(file_path_credentials, 'r') as file:
#****
# чтение данных из Excel
    df = pd.read_excel(file_path_credentials)
#df = pd.read_excel('C:/Users/maxim.rachinsky/Code/Testing/Input/data_in_v2.xlsx')  # путь к файлу данных
# названия столбцов с данными
df.columns = ['name', 'site1', 's1', 'd1', 'tip1', 'h1', 'r1', 'site2', 's2', 'd2', 'h2', 'r2', 'distan']

# создание KML-файла
kml = simplekml.Kml()
bs_folder = kml.newfolder(name='Существующий')  # папка БС
kandidat_folder = kml.newfolder(name='Кандидат')  # папка кондидата
lines_folder = kml.newfolder(name='Пролет')  # папка РРЛ

# Создание множеств для записей точек и линий
points = set()
lines = set()

# создание линий и точек на основе координат из Excel
for index, row in df.iterrows():
    dist = geodesic((row['s1'], row['d1']), (row['s2'], row['d2'])).km  # расчет растояния
    distance = round(dist, 2)

    point1 = f"{row['s1']},{row['d1']},{row['h1']}"
    point2 = f"{row['s2']},{row['d2']},{row['h2']}"
    line = (point1, point2)

    # Проверка, есть ли уже такая точка и линия
    if point1 not in points:
        point1_coords = [row['d1'], row['s1'], row['h1']]
        kandidat_point = kandidat_folder.newpoint(name=row['site1'], coords=[point1_coords])
        kandidat_point.description = f"Широта/Долгота : {row['s1']} / {row['d1']}"
        kandidat_point.style.iconstyle.color = 'fff5ef42'

        points.add(point1)
    if point2 not in points:
        point2_coords = [row['d2'], row['s2'], row['h2']]
        bs_point = bs_folder.newpoint(name=row['site2'], coords=[point2_coords])
        bs_point.description = f"Широта/Долгота : {row['s2']} / {row['d2']}"
        bs_point.style.iconstyle.color = 'ff493fb0'

        points.add(point2)
    if line not in lines:
        line_obj = lines_folder.newlinestring(name=row['name'],
                                              coords=[(row['d1'], row['s1'], row['h1']),
                                                      (row['d2'], row['s2'], row['h2'])])
        #line_obj.description = f"Пролет: {row['name']}"
        #line_obj.description = f"Тип сайта: {row['tip1']}"
        #line_obj.description = f"Высота подвеса: {row['h1']}/{row['h2']}"
        #line_obj.description = f"Диаметр антен: {row['r1']}/{row['r2']}"
        line_obj.extendeddata.newdata(name='Пролет: ', value=row['name'])
        line_obj.extendeddata.newdata(name='Тип сайта: ', value=row['tip1'])
        line_obj.extendeddata.newdata(name='Высота подвеса: ', value=f"{row['h1']}/{row['h2']}")
        line_obj.extendeddata.newdata(name='Диаметр антен: ', value=f"{row['r1']}/{row['r2']}")
        line_obj.extendeddata.newdata(name="Distance: ", value=f"{distance} km")  # добавление расстояния к точкам
        line_obj.style.linestyle.color = 'fff5ef42'  # цвет линии
        line_obj.style.linestyle.width = 3  # ширина линии
        lines.add(line)

kml.save('project.kml')  # путь KML-файла
#kml.save('C:/Users/maxim.rachinsky/Code/Testing/Input/project.kml')  # путь KML-файлаok