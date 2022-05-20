import pandas as pd
import numpy as np
from simpledbf import Dbf5

path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\1nuom.DBF"

dbf = Dbf5(path, codec='cp1251')

df = dbf.to_dataframe()

z = df.groupby('FEA_CODE').size()

WELDS_ID = [901, 902, 903, 904, 905, 906]
WELDS_NAMES = ['1', '2', '3', '4', '5', '6']

z1 = df[df.FEA_CODE.isin(WELDS_ID)]
z2 = z1[['FEA_CODE', 'WT']]

z2.loc[z2["FEA_CODE"] == 904, "FEA_CODE"] = "Бесшовная"

print(z2)
