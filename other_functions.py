import pandas as pd

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
        fea_code_replace_ft = str(fea_code_replace_ft).lower()
    if har_code1 in ft_list:
        har_code1_ft = str(empty_if_nan(har_code1)).lower()
    if har_code2 in ft_list:
        har_code2_ft = str(empty_if_nan(har_code2)).lower()

    corr = str(empty_if_nan(corr)).lower()
    description = str(empty_if_nan(description)).lower()

    feature_type = f"{fea_code_replace_ft}, {corr}, {har_code1_ft}, {har_code2_ft}"

    feature_type = feature_type.replace(', , , , ', ', , , ')
    feature_type = feature_type.replace(', , , ', ', , ')
    feature_type = feature_type.replace(', , ', ', ')
    if feature_type[:2] == ', ':
        feature_type = feature_type[2:]

    if feature_type[-2:] == ', ':
        feature_type = feature_type[:-2]

    return str(feature_type).capitalize()


def dbf_description_combine(har_code1, har_code2, corr, description, ft_list, d_list):

    har_code1_d = ''
    har_code2_d = ''

    if har_code1 in d_list:
        har_code1_d = str(empty_if_nan(har_code1)).lower()
    if har_code2 in d_list:
        har_code2_d = str(empty_if_nan(har_code2)).lower()

    corr = str(empty_if_nan(corr)).lower()
    description = str(empty_if_nan(description)).lower()

    description_return = f"{har_code1_d}, {har_code2_d}, {description}"

    description_return = description_return.replace(', , , ', ', , ')
    description_return = description_return.replace(', , ', ', ')
    if description_return[:2] == ', ':
        description_return = description_return[2:]

    if description_return[-2:] == ', ':
        description_return = description_return[:-2]

    return str(description_return).capitalize()
