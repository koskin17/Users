import pandas as pd


df = pd.ExcelFile('users.xls', engine='xlrd')

print(df)
