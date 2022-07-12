path = r'd:\WORK\#Thailand\NXE inch SRC to LLK, 134 km\Reports\FR'

float_inches = {'4': 4.5,
                '5': 5.563,
                '6': 6.625,
                '8': 8.625,
                '10': 10.75,
                '12': 12.75}


def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False


def parse_inch(file_path):
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
        print("Parse inch Error: inch found but No Diam /", ex)
        return None

    if isfloat(inv_inch):
        if inv_inch in float_inches:
            return float_inches[inv_inch]
        return inv_inch
    else:
        return None


if __name__ == '__main__':

    path = r"z:\Projects\Russia\Orenburgneft\NOD 10 inch ННП ДНС Рыбкинская-УПН Загорская 2 й участок, 24.5 km\Reports\FR\2021.12\Database"
    inch = parse_inch(path)
    print(inch)
