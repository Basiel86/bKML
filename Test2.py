import pandas as pd

if __name__ == '__main__':
    df1 = pd.DataFrame({'lkey': ['foo', 'bar', 'baz', 'foo5'],
                        'value': [1, 2, 3, 5]})
    df2 = pd.DataFrame({'lkey': ['foo', 'bar', 'baz', 'foo1'],
                        'value': [2, 2, 3, 6]})

    df1=pd.concat([df1, df2]).drop_duplicates().reset_index(drop=True)

    print(df1)
