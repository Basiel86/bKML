import pandas as pd

if __name__ == '__main__':
    # Creating dataframe
    df1 = pd.DataFrame({'SECT_NUM': ['10', '10',
                                     '20', '30'],
                        'JL': [0, 0, 0, 0]})

    df2 = pd.DataFrame({'SECT_NUM': ['10', '20',
                                     '30', '40',
                                     '50'],
                        'JL': [1, 2, 3, 4, 5]})

    print(df1)
    print(df2)

    df1['JL'] = df1['SECT_NUM'].map(df2.set_index('SECT_NUM')['JL'])

    print(df1)
