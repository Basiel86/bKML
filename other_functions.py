import pyfiglet
import os


def print_all_pyfiglet():
    fig = pyfiglet.Figlet()
    for font in fig.getFonts():
        print(font + '\n\n' + pyfiglet.figlet_format('DBF to KML', font=font), '\n\n')


def cls_pyfiglet(pyfiglet_msg, pyfiglet_font):
    os.system('cls')
    # 'fender' 'kban' 'larry3d'
    print('\n' + pyfiglet.figlet_format(pyfiglet_msg, justify='center', font=pyfiglet_font))



def empty_if_nan(poss_nan):
    if str(poss_nan) == 'nan':
        return ''
    else:
        return poss_nan


def dbf_feature_type_combine(fea_code_replace_ft, har_code1, har_code2, corr, description, ft_list, d_list):
    fea_code_replace_ft = empty_if_nan(fea_code_replace_ft)
    har_code1_ft = ''
    har_code2_ft = ''
    if fea_code_replace_ft in ft_list:
        fea_code_replace_ft = fea_code_replace_ft
    if har_code1 in ft_list or isinstance(har_code1, int):
        har_code1_ft = empty_if_nan(har_code1)
    if har_code2 in ft_list or isinstance(har_code2, int):
        har_code2_ft = empty_if_nan(har_code2)

    corr = empty_if_nan(corr)
    description = empty_if_nan(description)

    feature_type = f"{fea_code_replace_ft}, {har_code2_ft}, {corr}, {har_code1_ft}"

    feature_type = feature_type.replace(', , , , ', ', , , ')
    feature_type = feature_type.replace(', , , ', ', , ')
    feature_type = feature_type.replace(', , ', ', ')
    if feature_type[:2] == ', ':
        feature_type = feature_type[2:]

    if feature_type[-2:] == ', ':
        feature_type = feature_type[:-2]

    if len(feature_type) > 1:
        feature_type = feature_type[0].upper() + feature_type[1:]

    return feature_type


def dbf_description_combine(har_code1, har_code2, corr, description, ft_list, d_list):
    har_code1_d = ''
    har_code2_d = ''

    if har_code2 == 65545:
        pass

    if har_code1 in d_list or isinstance(har_code1, int):
        har_code1_d = empty_if_nan(har_code1)
    if har_code2 in d_list or isinstance(har_code2, int):
        har_code2_d = empty_if_nan(har_code2)

    corr = empty_if_nan(corr)
    description = empty_if_nan(description)

    description_return = f"{har_code1_d}, {har_code2_d}, {description}"

    description_return = description_return.replace(', , , ', ', , ')
    description_return = description_return.replace(', , ', ', ')
    if description_return[:2] == ', ':
        description_return = description_return[2:]

    if description_return[-2:] == ', ':
        description_return = description_return[:-2]

    if len(description_return) > 0:
        description_return = description_return[0].upper() + description_return[1:]

    return description_return


if __name__ == '__main__':
    print_all_pyfiglet()
    input("Press Any Key")
