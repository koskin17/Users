import os
from datetime import datetime

import pandas as pd

print("Загрузка данных по пользователям...")

"""Load data about users"""
df_users = pd.read_excel('user_admin.xlsx',
                         na_values="NA",
                         usecols=['ID', 'Баллы', 'Последняя авторизация в приложении',
                                  'Страна', 'Город работы', 'Тип пользователя',
                                  'Фамилия', 'Имя', 'Отчество', 'E-Mail'],
                         converters={"ID": int, "Баллы": int})
df_users['Последняя авторизация в приложении'] = pd.to_datetime(df_users['Последняя авторизация в приложении'],
                                                                format='%d.%m.%Y %H:%M:%S').dt.normalize()
df_users['Баллы'].fillna(0, inplace=True)
df_users = df_users.fillna('')

"""Columns for check data about users"""
df_users_columns = ['ID',
                    'Баллы',
                    'Последняя авторизация в приложении',
                    'Город работы',
                    'Страна',
                    'Тип пользователя',
                    'Фамилия',
                    'Имя',
                    'Отчество',
                    'E-Mail']


def check_file_with_users():
    """Check the availability necessary columns in file about users"""

    for col_name in df_users_columns:
        if col_name not in df_users.columns:
            print(f"В загруженных данных не хватает столбца {col_name}")
            return False

    return True


"""Clean spam and test accounts in Users_DataFrame"""
df_users = df_users[(df_users['Страна'] != '') &
                    (df_users['Тип пользователя'] != 'Клиент')]  # exception empty row in "Страна" and 'Клиент' as spam
print("Данные по пользователям загружены.")

exclude_users = ['kazah89', 'sanin, ''samoilov', 'axorindustry', 'kreknina', 'zeykin', 'berdnikova', 'ostashenko',
                 'skalar', 'test', 'malyigor', 'ihormaly', 'axor', 'kosits']  # sings of exclude account from counting

"""Creating list of excluded accounts"""
exclude_list = set()  # list for exclude accounts from count: test, axor and so on
for email in df_users['E-Mail']:
    for i in exclude_users:
        if i in email:
            exclude_list.add(email)

"""Clean DataFrame from exclude accounts"""
df_users = df_users.loc[~df_users['E-Mail'].isin(exclude_list)]
print('Список исключаемых аккаунтов сформирован.')

countries = list(set(df_users["Страна"]))  # list of countries in DataFrame

"""Columns for check data about scans"""
df_scans_columns = ['UF_TYPE',
                    'UF_POINTS',
                    'Дилер+Монтажник',
                    'UF_USER_ID',
                    'Монтажник',
                    'UF_CREATED_AT',
                    'Страна',
                    'Сам себе',
                    'Монтажник.1']

print("Загрузка данных по сканам...")
df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx',
                         usecols=['UF_TYPE', 'UF_POINTS', 'Дилер+Монтажник', 'UF_USER_ID',
                                  'Монтажник', 'UF_CREATED_AT', 'Страна', 'Сам себе',
                                  'Монтажник.1'],
                         converters={"UF_POINTS": int, "UF_USER_ID": int, "Монтажник": int})

df_scans['UF_CREATED_AT'] = pd.to_datetime(df_scans['UF_CREATED_AT'], format='%d.%m.%Y %H:%M:%S').dt.normalize()
df_scans = df_scans.fillna('')


def check_file_with_scans():
    """Check the availability necessary columns in file about scans"""

    for col_name in df_scans_columns:
        if col_name not in df_scans.columns:
            print(f"В загруженных данных не хватает столбца {col_name}")
            return False

    print("Данные по сканам загружены.")

    return True


def users_by_country():
    """Formation general statistics about users by countries."""

    list_for_df = []
    for country in countries:
        tmp_list = [country]
        data = df_users[(df_users["Страна"] == country) &
                        (df_users["Тип пользователя"] == "Дилер")]
        tmp_list.append(len(data["ID"]))
        data = df_users[(df_users["Страна"] == country) &
                        (df_users["Тип пользователя"] == "Монтажник")]
        tmp_list.append(len(data["ID"]))
        tmp_list.insert(1, sum(tmp_list[1:]))
        list_for_df.append(tmp_list)

    columns_for_df = ['Страна', 'Всего пользователей', 'Дилеров', 'Монтажников']
    index_for_df = range(len(list_for_df))
    total_stat_df = pd.DataFrame(list_for_df, index_for_df, columns_for_df).sort_values(by="Всего пользователей",
                                                                                        ascending=False)
    total_stat_df.to_excel(f"total_stats_about_users_for_{datetime.now().date()}.xlsx")
    os.startfile(f'total_stats_about_users_for_{datetime.now().date()}.xlsx')


def last_authorization_in_app():
    """ Information about last authorisation users in app by years"""

    df_users['Year'] = df_users['Последняя авторизация в приложении'].dt.year
    df_users['Year'].fillna(0, inplace=True)
    # years = sorted(map(int, [year for year in set(df_users['Year']) if year != '']))
    years = sorted([year for year in set(df_users['Year']) if year != 0])

    def amount_users_by_type(country_for_amount_users: str, user_type: str):
        """ Count amount users in country. """

        data = df_users[(df_users["Страна"] == country_for_amount_users) & (df_users["Тип пользователя"] == user_type)]
        amount_of_users = len(data["ID"])

        return amount_of_users

    def last_authorization(year: int, user_type: str, country_for_last_authorization: str):
        """Counting quantity of users with last authorisation in specific year."""

        if year == 0:
            data = df_users[
                (df_users['Year'] == 0) &
                (df_users['Тип пользователя'] == user_type) &
                (df_users['Страна'] == country_for_last_authorization)]

            last = len(data["ID"])
        else:
            data = df_users[(df_users['Year'] == year) &
                            (df_users["Тип пользователя"] == user_type) &
                            (df_users["Страна"] == country_for_last_authorization)]

            last = len(data["ID"])

        return last

    last_authorization_in_app_list = []

    for country in countries:
        last_authorization_in_app_list.append([country, 'Дилеры', amount_users_by_type(country, 'Дилер')] +
                                              [last_authorization(year, 'Дилер', country) for year in years] +
                                              [last_authorization(0, 'Дилер', country)])

    last_authorization_in_app_list.append(['', '', '', '', '', '', '', '', ])

    for country in countries:
        last_authorization_in_app_list.append([country, 'Монтажники', amount_users_by_type(country, 'Монтажник')] +
                                              [last_authorization(year, 'Монтажник', country) for year in years] +
                                              [last_authorization(0, 'Монтажник', country)])

    columns = ['Страна', 'Тип пользователей', 'Всего в базе'] + [year for year in years] + ['Не авторизировались']
    index = [_ for _ in range(len(last_authorization_in_app_list))]
    last_authorization_in_app_df = pd.DataFrame(last_authorization_in_app_list, index, columns)

    last_authorization_in_app_df.to_excel(f'last_authorization_in_app {datetime.now().date()}.xlsx')
    os.startfile(f'last_authorization_in_app {datetime.now().date()}.xlsx')


def authorization_during_period(start_date, end_date):
    """ information about the amount of authorized users for the period """

    def period_data(start_period_of_authorisation: datetime, end_period_of_authorization: datetime, user_type: str,
                    authorization_in_country: str):
        """ Counting the amount of users authorized in App during period """

        data = df_users[(df_users['Тип пользователя'] == user_type) &
                        (df_users['Страна'] == authorization_in_country) &
                        (df_users['Последняя авторизация в приложении'] >= start_period_of_authorisation) &
                        (df_users['Последняя авторизация в приложении'] <= end_period_of_authorization)]

        authorization = len(data['ID'])

        return authorization

    total_amount = 0
    authorization_during_period_list = []
    for country in countries:
        amount_of_dealers = period_data(start_date, end_date, 'Дилер', country)
        authorization_during_period_list.append([country, 'Дилеры', amount_of_dealers])
        total_amount += amount_of_dealers
        amount_of_adjusters = period_data(start_date, end_date, 'Монтажник', country)
        authorization_during_period_list.append([country, 'Монтажники', amount_of_adjusters])
        total_amount += amount_of_adjusters
        authorization_during_period_list.append(['', '', ''])

    authorization_during_period_list.append(['Всего:', '', total_amount])

    columns = ['Страна', 'Тип пользователей', 'Авторизовалось пользователей']
    index = [_ for _ in range(len(authorization_during_period_list))]
    authorization_during_period_df = pd.DataFrame(authorization_during_period_list, index, columns)

    start = datetime.strftime(start_date, "%d-%m-%Y")
    end = datetime.strftime(end_date, "%d-%m-%Y")

    authorization_during_period_df.to_excel(f"authorization_during_period {start}-{end}.xlsx")
    os.startfile(f'authorization_during_period {start}-{end}.xlsx')


def points_by_users_and_countries():
    """ Information about points by users and countries """

    def sum_of_points(type_of_user: str, country_for_sum_points: str):
        """ Count point of users by country"""

        data = df_users[(df_users['Тип пользователя'] == type_of_user) &
                        (df_users['Страна'] == country_for_sum_points)]

        points_of_users_in_country = sum(data['Баллы'])

        return points_of_users_in_country

    points_by_users_and_countries_list = []
    for country in countries:
        total_points = 0
        points = sum_of_points('Дилер', country)
        points_by_users_and_countries_list.append([country, 'Дилеры', points])
        total_points += points
        points = sum_of_points('Монтажник', country)
        points_by_users_and_countries_list.append([country, 'Монтажники', points])
        total_points += points
        points_by_users_and_countries_list.append(['Всего баллов:', '', total_points])
        points_by_users_and_countries_list.append(['', '', ''])

    columns = ['Страна', 'Тип пользователей', 'Сумма баллов']
    index = [_ for _ in range(len(points_by_users_and_countries_list))]
    points_by_users_and_countries_df = pd.DataFrame(points_by_users_and_countries_list, index, columns)

    points_by_users_and_countries_df.to_excel(f"points_by_users_and_countries {datetime.now().date()}.xlsx")
    os.startfile(f'points_by_users_and_countries {datetime.now().date()}.xlsx')


def data_about_scan_users_in_current_year():
    """ Information about scanned users in current year """

    def scanned_users(country_for_scanned_users: str, user_type: str, himself=True):
        """ Count amount of users scanned in current year"""

        count = set()

        if himself:
            if user_type == 'Дилер':
                data = df_scans[(df_scans['Страна'] == country_for_scanned_users) &
                                (df_scans['Сам себе'] == user_type) &
                                (df_scans['Монтажник.1'] == '')]

                count = set(data['UF_USER_ID'])

            elif user_type == 'Монтажник':
                data = df_scans[(df_scans['Страна'] == country_for_scanned_users) &
                                (df_scans['Сам себе'] == 'Монтажник')]

                count = set(data['UF_USER_ID'])

        else:
            data = df_scans[(df_scans['Страна'] == country_for_scanned_users) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

            count = set(data['Монтажник'])

        return len(count)

    table_about_scan_users_in_year_list = []
    for country in countries:
        dealers_himself = scanned_users(country, 'Дилер')
        adjusters_himself = scanned_users(country, 'Монтажник')
        adjusters_for_dealers = scanned_users(country, 'Монтажник', False)
        table_about_scan_users_in_year_list.append([country, 'Дилеры', 'Сами себе', dealers_himself])
        table_about_scan_users_in_year_list.append(['', 'Монтажники', 'Сами себе', adjusters_himself])
        table_about_scan_users_in_year_list.append(['', 'Монтажники', 'Сканировали дилеру', adjusters_for_dealers])
        table_about_scan_users_in_year_list.append(['', '', 'Итого:',
                                                    dealers_himself + adjusters_himself + adjusters_for_dealers])
        table_about_scan_users_in_year_list.append(['', '', '', ''])

    columns = ['Страна', 'Тип пользователей', 'Сканировали', 'Кол-во пользователей']
    index = [_ for _ in range(len(table_about_scan_users_in_year_list))]
    table_about_scan_users_in_year_df = pd.DataFrame(table_about_scan_users_in_year_list, index, columns)

    table_about_scan_users_in_year_df.to_excel(f"scanned_users_in_year {datetime.now().date()}.xlsx")
    os.startfile(f'scanned_users_in_year {datetime.now().date()}.xlsx')


def data_about_points():
    """ Information about sum of points scanned in current year """

    def total_amount_of_points_for_year(country_for_points, user_type):
        """Count the sum of points scanned in current year"""

        amount_of_points = 0

        if user_type == 'Дилер':
            data = df_scans[(df_scans['Страна'] == country_for_points) &
                            (df_scans['Сам себе'] == user_type)]

            amount_of_points = sum(data['UF_POINTS'])

        elif user_type == 'Монтажник':
            data = df_scans[(df_scans['Страна'] == country_for_points) &
                            (df_scans['Сам себе'] == user_type)]

            amount_of_points = sum(data['UF_POINTS'])

            data = df_scans[(df_scans['Страна'] == country_for_points) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

            amount_of_points += sum(data['UF_POINTS'])

        return amount_of_points

    """ Output data about points """
    data_about_points_lst = []
    for country in countries:
        point_of_dealers = total_amount_of_points_for_year(country, 'Дилер')
        points_of_adjusters = total_amount_of_points_for_year(country, 'Монтажник')
        data_about_points_lst.append([country, 'Дилеры', point_of_dealers])
        data_about_points_lst.append(['', 'Монтажники', points_of_adjusters])
        data_about_points_lst.append(['', 'Итого:', point_of_dealers + points_of_adjusters])
        data_about_points_lst.append(['', '', ''])

    columns = ['Страна', 'Тип пользователей', 'Насканировано баллов']
    index = [_ for _ in range(len(data_about_points_lst))]
    data_about_points_df = pd.DataFrame(data_about_points_lst, index, columns)

    data_about_points_df.to_excel(f"all_points_of_users_by_country {datetime.now().date()}.xlsx")
    os.startfile(f"all_points_of_users_by_country {datetime.now().date()}.xlsx")


def top_users_by_scans(country: str, user_type: str):
    """ TOP dealers / adjusters by scans"""
    top_users = {}  # dictionary for TOP users by points
    top_users_by_scans_lst = []
    surname = {}  # dictionary for surnames

    """ Filling the dictionary of surnames"""
    for df_users_ID, df_users_surname in zip(df_users['ID'], df_users['Фамилия']):
        surname[df_users_ID] = df_users_surname

    if user_type == 'Дилер':
        data = df_scans[(df_scans['Страна'] == country) &
                        (df_scans['Сам себе'] == user_type)]

        for df_scans_dealer_id, df_scans_points in zip(data['UF_USER_ID'], data['UF_POINTS']):
            if df_scans_dealer_id in top_users.keys():
                top_users[df_scans_dealer_id] += df_scans_points
            else:
                top_users[df_scans_dealer_id] = df_scans_points

        for df_scans_dealer_id in top_users.keys():
            if df_scans_dealer_id in surname.keys():  # some users don't fill "Страна" and they don't count in df_users
                top_users_by_scans_lst.append([df_scans_dealer_id,
                                               surname[df_scans_dealer_id],
                                               top_users[df_scans_dealer_id]])

        top_users_by_scans_lst = sorted(top_users_by_scans_lst, key=lambda x: x[2], reverse=True)

        top_users_by_scans_lst.append(['Итого:', '', sum(top_users.values())])

        columns = ['ID пользователя', 'Фамилия', 'Сумма насканированных баллов']
        index = [_ for _ in range(len(top_users_by_scans_lst))]
        top_users_by_scans_list_df = pd.DataFrame(top_users_by_scans_lst, index, columns)

        top_users_by_scans_list_df.to_excel(f"TOP_dealers_by_scans_in_{country} {datetime.now().date()}.xlsx")
        os.startfile(f"TOP_dealers_by_scans_in_{country} {datetime.now().date()}.xlsx")

    elif user_type == 'Монтажник':
        data = df_scans[(df_scans['Страна'] == country) &
                        (df_scans['Сам себе'] == user_type)]

        for df_scans_adjuster_id, df_scans_point in zip(data['UF_USER_ID'], data['UF_POINTS']):
            if df_scans_adjuster_id in top_users.keys():
                top_users[df_scans_adjuster_id] += df_scans_point
            else:
                top_users[df_scans_adjuster_id] = df_scans_point

        data = df_scans[(df_scans['Страна'] == country) &
                        (df_scans['Монтажник.1'] == user_type)]

        for df_scans_adjuster_id, df_scans_point in zip(data['Монтажник'], data['UF_POINTS']):
            if df_scans_adjuster_id in top_users.keys():
                top_users[df_scans_adjuster_id] += df_scans_point
            else:
                top_users[df_scans_adjuster_id] = df_scans_point

        for df_scans_adjuster_id in top_users.keys():
            if df_scans_adjuster_id in surname.keys():  # some users don't fill "Страна" and they not count in df_users
                top_users_by_scans_lst.append([df_scans_adjuster_id,
                                               surname[df_scans_adjuster_id],
                                               top_users[df_scans_adjuster_id]])

        top_users_by_scans_lst = sorted(top_users_by_scans_lst, key=lambda x: x[2], reverse=True)

        top_users_by_scans_lst.append(['Итого:', '', sum(top_users.values())])

        columns = ['ID пользователя', 'Фамилия', 'Сумма насканированных баллов']
        index = [_ for _ in range(len(top_users_by_scans_lst))]
        top_users_by_scans_list_df = pd.DataFrame(top_users_by_scans_lst, index, columns)

        top_users_by_scans_list_df.to_excel(f"TOP_adjusters_by_scans_in_{country} {datetime.now().date()}.xlsx")
        os.startfile(f"TOP_adjusters_by_scans_in_{country} {datetime.now().date()}.xlsx")


def data_about_scans_during_period(start_date: datetime, end_date: datetime):
    """Output information about users and scans during period"""

    def scanned_users_per_period(country_for_scanned_users_in_period: str, user_type: str,
                                 start_period_of_scanned_users: datetime,
                                 end_period_of_scanned_users: datetime, himself=True):
        """ Count amount of users scanned during period"""

        count = set()
        if himself:
            if user_type == 'Дилер':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_of_scanned_users) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_scanned_users) &
                                (df_scans['Страна'] == country_for_scanned_users_in_period) &
                                (df_scans['Сам себе'] == user_type) &
                                (df_scans['Монтажник.1'] != 'Монтажник')]

                count = set(data['UF_USER_ID'])

            elif user_type == 'Монтажник':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_of_scanned_users) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_scanned_users) &
                                (df_scans['Страна'] == country_for_scanned_users_in_period) &
                                (df_scans['Сам себе'] == user_type)]

                count = set(data['UF_USER_ID'])
        else:
            for date_of_scan, df_ID, df_country, df_adjuster_1 in zip(df_scans['UF_CREATED_AT'], df_scans['Монтажник'],
                                                                      df_scans['Страна'], df_scans['Монтажник.1']):

                if start_period_of_scanned_users <= date_of_scan <= end_period_of_scanned_users:
                    if country == df_country and df_adjuster_1 == 'Монтажник':
                        count.add(df_ID)

        return len(count)

    def sum_of_points_per_period(country_for_sum_points_in_period: str, user_type: str,
                                 start_period_for_sum_points: datetime, end_period_of_sum_points: datetime,
                                 himself=True):
        """ Count sum of scanned points during period """

        amount_of_points_per_period = 0

        if himself:
            if user_type == 'Дилер':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                                (df_scans['Страна'] == country_for_sum_points_in_period) &
                                (df_scans['Сам себе'] == user_type) &
                                (df_scans['Монтажник.1'] != 'Монтажник')]

                amount_of_points_per_period = sum(data['UF_POINTS'])

            elif user_type == 'Монтажник':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                                (df_scans['Страна'] == country_for_sum_points_in_period) &
                                (df_scans['Сам себе'] == user_type)]

                amount_of_points_per_period = sum(data['UF_POINTS'])

        else:
            data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                            (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                            (df_scans['Страна'] == country_for_sum_points_in_period) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

            amount_of_points_per_period = sum(data['UF_POINTS'])

        return amount_of_points_per_period

    """Dates of previous period"""
    time_delta = end_date - start_date
    previous_start = start_date - time_delta
    previous_end = end_date - time_delta

    data_about_scans_during_period_list = []

    for country in countries:
        data_about_scans_during_period_list.append([country,
                                                    'Дилеры',
                                                    'Сами себе',
                                                    scanned_users_per_period(country, 'Дилер', start_date, end_date,
                                                                             True),
                                                    scanned_users_per_period(country, 'Дилер', previous_start,
                                                                             previous_end, himself=True),
                                                    sum_of_points_per_period(country, 'Дилер', start_date,
                                                                             end_date) + sum_of_points_per_period(
                                                        country, 'Монтажник', start_date, end_date, himself=False)])
        data_about_scans_during_period_list.append(['',
                                                    'Монтажники',
                                                    'Сами себе',
                                                    scanned_users_per_period(country, 'Монтажник', start_date, end_date,
                                                                             True),
                                                    scanned_users_per_period(country, 'Монтажник', previous_start,
                                                                             previous_end, himself=True),
                                                    sum_of_points_per_period(country, 'Монтажник', start_date,
                                                                             end_date)])
        data_about_scans_during_period_list.append(['',
                                                    'Монтажники',
                                                    'Сканировали дилеру',
                                                    scanned_users_per_period(country, 'Монтажник', start_date, end_date,
                                                                             False),
                                                    scanned_users_per_period(country, 'Монтажник', previous_start,
                                                                             previous_end, False),
                                                    sum_of_points_per_period(country, 'Монтажник', start_date, end_date,
                                                                             himself=False)])
        data_about_scans_during_period_list.append(['',
                                                    '',
                                                    'Итого:',
                                                    scanned_users_per_period(country, 'Дилер', start_date, end_date,
                                                                             True) +
                                                    scanned_users_per_period(country, 'Монтажник', start_date, end_date,
                                                                             True) +
                                                    scanned_users_per_period(country, 'Монтажник', start_date, end_date,
                                                                             False),
                                                    scanned_users_per_period(country, 'Дилер', previous_start,
                                                                             previous_end, True) +
                                                    scanned_users_per_period(country, 'Монтажник', previous_start,
                                                                             previous_end, True) +
                                                    scanned_users_per_period(country, 'Монтажник', previous_start,
                                                                             previous_end, False),
                                                    sum_of_points_per_period(country, 'Дилер', start_date,
                                                                             end_date) + sum_of_points_per_period(
                                                        country, 'Монтажник', start_date, end_date, himself=False) +
                                                    sum_of_points_per_period(country, 'Монтажник', start_date,
                                                                             end_date) +
                                                    sum_of_points_per_period(country, 'Монтажник', start_date, end_date,
                                                                             himself=False)])
        data_about_scans_during_period_list.append(['',
                                                    '',
                                                    '',
                                                    '',
                                                    '',
                                                    ''])

    columns = ['Страна',
               'Пользователи',
               'Сканировали:',
               'Кол-во пользователей',
               'Пользователей за предыдущий аналогичный период',
               'Баллов за указанный период']
    index = [_ for _ in range(len(data_about_scans_during_period_list))]
    data_about_scans_during_period_df = pd.DataFrame(data_about_scans_during_period_list, index, columns)

    start = datetime.strftime(start_date, "%d-%m-%Y")
    end = datetime.strftime(end_date, "%d-%m-%Y")

    data_about_scans_during_period_df.to_excel(f"data_about_scans_during_period_{start}-{end}.xlsx")
    os.startfile(f'data_about_scans_during_period_{start}-{end}.xlsx')


def scanned_users_by_months():
    """ Information about scanned users by country in each month """
    months = {1: 'Январь',
              2: 'Февраль',
              3: 'Март',
              4: 'Апрель',
              5: 'Май',
              6: 'Июнь',
              7: 'Июль',
              8: 'Август',
              9: 'Сентябрь',
              10: 'Октябрь',
              11: 'Ноябрь',
              12: 'Декабрь'}
    df_scans['Месяц'] = df_scans['UF_CREATED_AT'].dt.month.map(months)

    scanned_users_by_months_list = []
    columns = ['Страна', 'Тип пользователей', 'Сканировали'] + [month for month in months.values()]

    for country in countries:
        dealers_himself = [country, 'Дилеры', 'Сами себе']
        adjusters_himself = ['', 'Монтажники', 'Сами себе']
        adjusters_for_dealers = ['', 'Монтажники', 'Сканировали дилеру']
        total_users_in_country = ['', '', 'Итого:']

        for month in months.values():
            data = df_scans[(df_scans['Страна'] == country) &
                            (df_scans['Месяц'] == month) &
                            (df_scans['Сам себе'] == 'Дилер') &
                            (df_scans['Монтажник.1'] == '')]

            d_h = len(set(data['UF_USER_ID']))  # dealers himself
            dealers_himself.append(d_h)

            data = df_scans[(df_scans['Месяц'] == month) &
                            (df_scans['Страна'] == country) &
                            (df_scans['Сам себе'] == 'Монтажник')]

            a_h = len(set(data['UF_USER_ID']))  # adjusters himself
            adjusters_himself.append(a_h)

            data = df_scans[(df_scans['Месяц'] == month) &
                            (df_scans['Страна'] == country) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

            a_d = len(set(data['Монтажник']))  # adjusters for dealers
            adjusters_for_dealers.append(a_d)

            total_users_in_country.append(d_h + a_h + a_d)

        scanned_users_by_months_list.append(dealers_himself)
        scanned_users_by_months_list.append(adjusters_himself)
        scanned_users_by_months_list.append(adjusters_for_dealers)
        scanned_users_by_months_list.append(total_users_in_country)
        scanned_users_by_months_list.append(['', '', ''])

    index = [_ for _ in range(len(scanned_users_by_months_list))]
    scanned_users_by_months_df = pd.DataFrame(scanned_users_by_months_list, index, columns)

    scanned_users_by_months_df.to_excel(f"scanned_users_by_months {datetime.now().date()}.xlsx")
    os.startfile(f"scanned_users_by_months {datetime.now().date()}.xlsx")


def finish():
    print('Всего хорошего!')
    pass
