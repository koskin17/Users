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
