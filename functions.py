import os
from datetime import datetime
import pandas as pd


def authorization_during_period(start_date, end_date):
    """ information about the amount of authorized users for the period """

    def period_data(start_period_of_authorisation: datetime, end_period_of_authorization: datetime, user_type: str,
                    authorization_in_country: str):
        """ Counting the amount of users authorized in App during period """

        data = df_users[(df_users['Тип пользователя'] == user_type) &
                        (df_users['Страна'] == authorization_in_country) &
                        (df_users['Последняя авторизация в приложении'] >= start_period_of_authorisation) &
                        (df_users['Последняя авторизация в приложении'] <= end_period_of_authorization)]

        return len(data['ID'])

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


# def data_about_points():
#     """ Information about sum of points scanned in current year """
#
#     def total_amount_of_points_for_year(country_for_points, user_type):
#         """Count the sum of points scanned in current year"""
#
#         amount_of_points = 0
#
#         if user_type == 'Дилер':
#             data = df_scans[(df_scans['Страна'] == country_for_points) &
#                             (df_scans['Сам себе'] == user_type)]
#
#             amount_of_points = sum(data['UF_POINTS'])
#
#         elif user_type == 'Монтажник':
#             data = df_scans[(df_scans['Страна'] == country_for_points) &
#                             (df_scans['Сам себе'] == user_type)]
#
#             amount_of_points = sum(data['UF_POINTS'])
#
#             data = df_scans[(df_scans['Страна'] == country_for_points) &
#                             (df_scans['Монтажник.1'] == 'Монтажник')]
#
#             amount_of_points += sum(data['UF_POINTS'])
#
#         return amount_of_points
#
#     """ Output data about points """
#     data_about_points_lst = []
#     for country in countries:
#         point_of_dealers = total_amount_of_points_for_year(country, 'Дилер')
#         points_of_adjusters = total_amount_of_points_for_year(country, 'Монтажник')
#         data_about_points_lst.append([country, 'Дилеры', point_of_dealers])
#         data_about_points_lst.append(['', 'Монтажники', points_of_adjusters])
#         data_about_points_lst.append(['', 'Итого:', point_of_dealers + points_of_adjusters])
#         data_about_points_lst.append(['', '', ''])
#
#     columns = ['Страна', 'Тип пользователей', 'Насканировано баллов']
#     index = [_ for _ in range(len(data_about_points_lst))]
#     data_about_points_df = pd.DataFrame(data_about_points_lst, index, columns)
#
#     data_about_points_df.to_excel(f"all_points_of_users_by_country {datetime.now().date()}.xlsx")
#     os.startfile(f"all_points_of_users_by_country {datetime.now().date()}.xlsx")


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
            data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_of_scanned_users) &
                            (df_scans['UF_CREATED_AT'] <= end_period_of_scanned_users) &
                            (df_scans['Страна'] == country_for_scanned_users_in_period) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

            count = set(data['Монтажник'])

        return len(count)

    def sum_of_points_per_period(country_for_sum_points_in_period: str, user_type: str,
                                 start_period_for_sum_points: datetime, end_period_of_sum_points: datetime,
                                 himself=True):
        """ Count sum of scanned points during period """

        if himself:
            if user_type == 'Дилер':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                                (df_scans['Страна'] == country_for_sum_points_in_period) &
                                (df_scans['Сам себе'] == user_type) &
                                (df_scans['Монтажник.1'] != 'Монтажник')]

            elif user_type == 'Монтажник':
                data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                                (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                                (df_scans['Страна'] == country_for_sum_points_in_period) &
                                (df_scans['Сам себе'] == user_type)]
        else:
            data = df_scans[(df_scans['UF_CREATED_AT'] >= start_period_for_sum_points) &
                            (df_scans['UF_CREATED_AT'] <= end_period_of_sum_points) &
                            (df_scans['Страна'] == country_for_sum_points_in_period) &
                            (df_scans['Монтажник.1'] == 'Монтажник')]

        return sum(data['UF_POINTS'])

    data_about_scans_during_period_list = []

    for country in countries:
        dealers_themselves = scanned_users_per_period(country, 'Дилер', start_date, end_date)
        adjusters_themselves = scanned_users_per_period(country, 'Монтажник', start_date, end_date)
        adjusters_for_dealers = scanned_users_per_period(country, 'Монтажник', start_date, end_date, False)
        points_of_dealers_themselves = sum_of_points_per_period(country, 'Дилер', start_date, end_date)
        points_of_adjusters_themselves = sum_of_points_per_period(country, 'Монтажник', start_date, end_date)
        points_of_adjusters_for_dealers = sum_of_points_per_period(country, 'Монтажник', start_date, end_date, False)

        data_about_scans_during_period_list.append([country, 'Дилеры', 'Сами себе', dealers_themselves,
                                                    points_of_dealers_themselves + points_of_adjusters_for_dealers])
        data_about_scans_during_period_list.append(['', 'Монтажники', 'Сами себе', adjusters_themselves,
                                                    points_of_adjusters_themselves])
        data_about_scans_during_period_list.append(['', 'Монтажники', 'Сканировали дилеру', adjusters_for_dealers,
                                                    points_of_adjusters_for_dealers])
        data_about_scans_during_period_list.append(['', '', 'Итого:', dealers_themselves + adjusters_themselves +
                                                    adjusters_for_dealers,
                                                    points_of_dealers_themselves + points_of_adjusters_for_dealers +
                                                    points_of_adjusters_themselves +
                                                    points_of_adjusters_for_dealers])
        data_about_scans_during_period_list.append(['', '', '', '', ''])

    columns = ['Страна',
               'Пользователи',
               'Сканировали:',
               'Кол-во пользователей',
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
