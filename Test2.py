import pandas as pd
import xlwings as xw
from pandasgui import show



if __name__ == '__main__':
    # Creating dataframe
    df1 = pd.DataFrame({'SECT_NUM': ['10', '11',
                                     '10', '10'],
                        'JL': [0, 0, 0, 0],
                        'DIST': [1, 2, 3, 4]})

    z = df1.loc[df1['SECT_NUM'] == '10']['SECT_NUM'].tolist()
    z2 = df1.loc[df1['SECT_NUM'] == '10']['DIST'].tolist()

    z3 = z + z2

    print(df1)

    # xw.view(df1, table=False)

    show(df1)




    df1.to_clipboard(index=False)
