import configparser
import os
import pandas as pd
import sys


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as ex:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


inch_list_path = resource_path(r'IDs\Inch_list.csv')

inch_mm_df = pd.read_csv(inch_list_path)
inch_names_list = inch_mm_df['Inch_name'].tolist()
inch_list = inch_mm_df['Inch'].tolist()
mm_list = inch_mm_df['MM'].tolist()
inch_dict = dict(zip(inch_names_list, inch_list))

float_inches = {'4': 4.5,
                '5': 5.563,
                '6': 6.625,
                '8': 8.625,
                '10': 10.75,
                '12': 12.75}


config = configparser.ConfigParser()


def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False


# парсинг инча из файла проектов (PRJ)
def parse_inch_prj(file_path):
    abspath = os.path.dirname(file_path)
    if os.path.exists(abspath):
        for filename in os.listdir(abspath):
            try:
                if ".prj" in filename:
                    config.read(os.path.join(abspath, filename))
                    # переменная с диаметром
                    diam_var = float(config['PARAMETERS']['PipeDiameter'])
                    # 1 - inch / 0 - мм
                    inch_or_mm = int(config['PARAMETERS']['DiameterDim'])

                    if inch_or_mm == 0:
                        diam_var_inch = round(diam_var / 25.4, 3)
                        result_inch = get_close_inch(diam_var_inch)
                        if diam_var_inch < result_inch:
                            print(
                                f"\n### Attention: No required Inch in InchList: PRJ={diam_var_inch} but Closest={result_inch} ###\n")
                    else:
                        result_inch = get_close_inch(diam_var)
                        # проверяем на корректность Инча в проекте
                        if diam_var not in inch_list:
                            print(
                                f"\n### Attention: Wrong Inch in PRJ file: PRJ={diam_var} but Closest={result_inch} ###\n")

                    print(f"# Inch parser (PRJ): {result_inch} inch")
                    return result_inch
            except KeyError:
                print(f"### ERROR Inch parser (PRJ): No PipeDiameter in PRJ file")
                return None
        return None


# возвращаем ближайший инч из списка
def get_close_inch(inch_val):
    closest_inch = min(inch_list, key=lambda x: abs(x - inch_val))
    return closest_inch


# парсинг инча из пути к файлу
def parse_inch_path(file_path):
    file_path = file_path.replace('  ', ' ')

    inch_loc = file_path.find('inch')
    if inch_loc == -1:
        return None

    inch_loc_inv = file_path[inch_loc - 2::-1]

    def ret_f_space(inch_loc_inv):
        for i in range(len(inch_loc_inv)):
            if inch_loc_inv[i] == " ":
                return inch_loc_inv[:i]

    inv_inch = ret_f_space(inch_loc_inv)
    try:
        inv_inch = inv_inch[::-1]
    except Exception as ex:
        # print("Parse inch Error: inch found but No Diam value /", ex)
        return None

    if isfloat(inv_inch):
        if inv_inch in float_inches:
            result_inch = get_close_inch(float_inches[inv_inch])
            print(f"# Inch parser (Path): {result_inch} inch")
            return result_inch

        result_inch = get_close_inch(float(inv_inch))
        print(f"# Inch parser (Path): {result_inch} inch")
        return result_inch
    else:
        return None


def get_inch_dict():
    return inch_dict


def get_inch_list():
    return inch_list


def get_inch_names_list():
    return inch_names_list


def parse_inch_combine(file_path):
    inch_of_path = parse_inch_path(file_path)
    inch_of_prj = parse_inch_prj(file_path)

    if inch_of_prj is None and inch_of_path is None:
        return None
    if inch_of_prj is None:
        print("#inch parser: Inch from Path")
        return inch_of_path
    else:
        print("#inch parser: Inch from PRJ")
        return inch_of_prj




if __name__ == '__main__':
    path = r"d:\WORK\SalymPetroleum\NWA 20 inch UPN-PSN, 88 km\Reports\FR\Database\1nwam.DBF"

    inch2 = parse_inch_prj(path)

    print(inch_list.index(inch2))


