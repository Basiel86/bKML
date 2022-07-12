import pandas as pd

if __name__ == '__main__':
    # Creating dataframe
    df1 = pd.DataFrame({'SECT_NUM': ['10', '11',
                                     '10', '10'],
                        'JL': [0, 0, 0, 0],
                        'DIST': [1, 2, 3, 4]})

    z = df1.loc[df1['SECT_NUM'] == '10']['SECT_NUM'].tolist()
    z2 = df1.loc[df1['SECT_NUM'] == '10']['DIST'].tolist()

    z3 = z + z2

    print(z)
    print(z2)
    print(z3)
