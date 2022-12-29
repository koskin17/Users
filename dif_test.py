import pandas as pd
from datetime import datetime

df_users = pd.read_excel('user_admin.xlsx')
df_users = df_users.fillna('')
# print(df_users.head())

# df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx')
# df_scans = df_scans.fillna('')
# print(df_scans.head())
countries = set(df_users["Страна"])
columns = ['ID',
                'Баллы',
                'Последняя авторизация в приложении',
                'Количество сеансов',
                'Город работы',
                'Страна',
                'Логин',
                'Тип пользователя',
                'Активность',
                'Дата регистрации',
                'Фамилия',
                'Имя',
                'Отчество',
                'E-Mail',
                'Последняя авторизация',
                'Дата изменения',
                'Город проживания',
                'Личный телефон',
                'Компания',
                'Назва дилера',
                'СПК 1',
                'СПК 2',
                'СПК 3',
                'СПК 4',
                'СПК 5']

print(columns == list(df_users.columns))

if all(name in df_users.columns for name in columns):
    print("Все есть")

