import pandas as pd

df_users = pd.read_excel('user_admin.xlsx')
print(df_users.head())
df_users = df_users.fillna('')
print(df_users.head())

df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx')
print(df_scans.head())
df_scans = df_scans.fillna('')
print(df_scans.head())