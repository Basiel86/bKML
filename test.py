import pandas as pd
import numpy as np


def same_as_upper(col: pd.Series) -> pd.Series:
    '''
    Recursively fill NaN rows with the previous value
    '''
    if any(pd.Series(col).isna()):
        col = pd.Series(np.where(col.isna(), col.shift(1), col))
        return same_as_upper(col)
    else:
        return col


if __name__ == '__main__':
    index_filename = ('test.xlsx')

    test = pd.read_excel(index_filename, sheet_name='Sheet1')

    test['JN'] = ""

    # rows = [i + 10 for i in range(0, len(welds)*10, 10)]
    # df = pd.Series(rows)
    # print(df)
    # test.loc[test['FEA_CODE_REPLACE'] == 'Шов', 'JN'] = df

    for i, row in test.iterrows():
        if test.loc[i]['FEA_CODE_REPLACE'] == 'Шов':
            test.at[i, 'JN'] = i

    # for i in range(0, 200, 10):
    #     print(i)

    test['JN'] = test['JN'].replace('', np.NAN)
    # test[test['JN'] == ""] = np.NaN
    test['JN'] = test['JN'].fillna(method='ffill')


    print(test)

    # welds = test[test['FEA_CODE_REPLACE'] == "Шов"][['FEA_DIST']]

    # считаем длину секций
    # test['JL'] = round(test['FEA_DIST'].diff(1), 3)
    # # двигаем длину секций на 1 вверх
    # test['JL'] = test['JL'].shift(-1)

    # print(test)
