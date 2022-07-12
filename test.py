import pandas as pd
import numpy as np

if __name__ == '__main__':

    feature_type = 'Коррозия, , , '

    feature_type = feature_type.replace(', , , ', ', , ')
    feature_type = feature_type.replace(', , ', ', ')
    if feature_type[:2] == ', ':
        feature_type = feature_type[2:]

    if feature_type[-2:] == ', ':
        feature_type = feature_type[:-2]


    print(feature_type)


