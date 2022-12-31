import pandas as pd
from pandas import to_datetime
import os

df_users = pd.read_excel('user_admin.xlsx',
                         converters={"Последняя авторизация в приложении": to_datetime})
df_users = df_users.fillna('')
# print(df_users.head())

# df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx')
# df_scans = df_scans.fillna('')
# print(df_scans.head())

exclude_users = ['kazah89', 'sanin, ''samoilov', 'axorindustry', 'kreknina', 'zeykin', 'berdnikova', 'ostashenko',
                 'skalar', 'test', 'malyigor', 'ihormaly', 'axor']
exclude_list = []

for email in df_users['E-Mail']:
    for i in exclude_users:
        if i in email and email not in exclude_list:
            exclude_list.append(email)

# print(exclude_list)
# df_users = df_users.loc[~df_users['E-Mail'].isin(exclude_list)]
# print(df_users)
df_datetime = df_users[2020 < df_users['Последняя авторизация в приложении'] < 2022]
df_datetime.to_excel('test.xlsx')
os.startfile('test.xlsx')
