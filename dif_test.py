import datetime
from datetime import datetime
import time

import pandas as pd
from pandas import to_datetime
import os

df_users = pd.read_excel('user_admin.xlsx',
                         na_values="NA",
                         converters={"ID": int, "Баллы": int})

#
# df_users['Авторизация в приложении'] = pd.to_datetime(df_users['Последняя авторизация в приложении'],
#                                                       format='%d.%m.%Y %H:%M:%S')
#
# df_users['Последняя авторизация в приложении'] = pd.to_datetime(df_users['Последняя авторизация в приложении'],
#                                                                 format='%d.%m.%Y %H:%M:%S').dt.normalize() #dt.date
# start = datetime.strptime(input("Дата начала в формате mm.dd.yyyy (через точку): "), '%d.%m.%Y')
# end = datetime.strptime(input("Дата конца в формате mm.dd.yyyy (через точку): "), '%d.%m.%Y')
# data = df_users[(start <= df_users['Последняя авторизация в приложении']) &
#                 (df_users['Последняя авторизация в приложении'] <= end)]
# data.to_excel('test.xlsx')

# print(start)
# end = datetime.strptime(input("Дата конца в формате mm.dd.yyyy (через точку): "), '%d.%m.%Y').date()
# print(end)


# data = df_users[(df_users['Авторизация в приложении'] >= datetime.date(2021, 7, 1)) &
#                  (df_users['Авторизация в приложении'] <= datetime.date(2021, 7, 31))]
#
# print(len(data['ID']))
# print(df_users['Авторизация в приложении'])
# for _ in df_users['Авторизация в приложении']:
#     if _ == datetime.date(2021, 6, 1):
#         print(_)
# print(df_users.dtypes)
# data = df_users[df_users['Авторизация'] == datetime.date(2021,6,1)]
# print(data)

# df_users['Баллы'].fillna(0, inplace=True)
# data = df_users[
#         (df_users['Тип пользователя'] == 'Дилер') &
#         (df_users['Страна'] == 'Украина')]
#
# point = sum(data['Баллы'])
# print(point)
# df_users['Month'] = df_users['Последняя авторизация в приложении'].dt.month
# print(df_users['Month'])
# df_users['Year'] = df_users['Последняя авторизация в приложении'].dt.year
# df_users.fillna('', inplace=True)
# print(df_users['Year'])
# data = df_users[df_users['Year'] == '']
# print(len(data["ID"]))
# print(len(data['ID']))
# df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx')
# df_scans = df_scans.fillna('')
# print(df_scans.head())

# exclude_users = ['kazah89', 'sanin, ''samoilov', 'axorindustry', 'kreknina', 'zeykin', 'berdnikova', 'ostashenko',
#                  'skalar', 'test', 'malyigor', 'ihormaly', 'axor']
# exclude_list = []
#
# for email in df_users['E-Mail']:
#     for i in exclude_users:
#         if i in email and email not in exclude_list:
#             exclude_list.append(email)

# print(exclude_list)
# df_users = df_users.loc[~df_users['E-Mail'].isin(exclude_list)]
# print(df_users)
# df_datetime = df_users[2020 < df_users['Последняя авторизация в приложении'] < 2022]
# df_datetime.to_excel('test.xlsx')
# os.startfile('test.xlsx')
