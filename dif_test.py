import datetime
from datetime import datetime
import time

import pandas as pd
from pandas import to_datetime
import os
import sys
from PyQt5.QtWidgets import *
from classes import *



# df_users = pd.read_excel('user_admin.xlsx',
#                          usecols=['ID', 'Баллы', 'Последняя авторизация в приложении',
#                                   'Страна', 'Тип пользователя', 'Фамилия', 'Имя',
#                                   'Отчество', 'E-Mail'],
#                          na_values="NA",
#                          converters={"ID": int, "Баллы": int})
#
# df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx',
#                          usecols=['UF_TYPE', 'UF_POINTS', 'Дилер+Монтажник', 'UF_USER_ID',
#                                   'Монтажник', 'UF_CREATED_AT', 'Страна', 'Сам себе',
#                                   'Монтажник.1'],
#                          converters={"UF_POINTS": int, "UF_USER_ID": int, "Монтажник": int})
#
# df_scans['UF_CREATED_AT'] = pd.to_datetime(df_scans['UF_CREATED_AT'], format='%d.%m.%Y %H:%M:%S').dt.normalize()

# top_users = {}  # словарь для ТОП пользователей
# surname = {}
# data = df_scans[(df_scans['Страна'] == 'Украина') &
#                 (df_scans['Сам себе'] == 'Дилер')]
#
# for df_users_ID, df_users_surname in zip(df_users['ID'], df_users['Фамилия']):
#     surname[df_users_ID] = df_users_surname
#
# for user_id, points in zip(data['UF_USER_ID'], data['UF_POINTS']):
#     if user_id in top_users.keys():
#         top_users[user_id] += points
#     else:
#         top_users[user_id] = points
#
# top_users_by_scans_lst = []
#
# for df_scans_user_id in top_users.keys():
#     top_users_by_scans_lst.append([df_scans_user_id, surname[df_scans_user_id], top_users[df_scans_user_id]])
#
# print(sorted(top_users_by_scans_lst, key=lambda x: x[2], reverse=True))

# months = {1: 'Январь',
#               2: 'Февраль',
#               3: 'Март',
#               4: 'Апрель',
#               5: 'Май',
#               6: 'Июнь',
#               7: 'Июль',
#               8: 'Август',
#               9: 'Сентябрь',
#               10: 'Октябрь',
#               11: 'Ноябрь',
#               12: 'Декабрь'}
# df_scans['Месяц'] = df_scans['UF_CREATED_AT'].dt.month.map(months)
# columns = ['Страна', 'Тип пользователей', 'Сканировали'] + [month for month in months.values()]
# print(columns)
#
# df_users['Последняя авторизация в приложении'] = pd.to_datetime(df_users['Последняя авторизация в приложении'],
#                                                                 format='%d.%m.%Y %H:%M:%S').dt.normalize()
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
