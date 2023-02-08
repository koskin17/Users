from PyQt5.QtWidgets import *
from functions import *


class MainWindows(QDialog):
    def __init__(self):
        super().__init__()
        # set Title for main windows
        self.setWindowTitle("Данные по пользователя и сканам в приложении AXOR")
        # # Set size of main window
        self.resize(600, 400)
        self.users_in_countries = QPushButton("Пользователи по странам", self)
        self.users_in_countries.move(0, 10)
        self.users_in_countries.clicked.connect(self.users_by_country)

        self.authorization_in_app = QPushButton("Авторизация пользователей в приложении", self)
        self.authorization_in_app.move(0, 45)
        self.authorization_in_app.clicked.connect(self.last_authorization_in_app)

        self.authorization_in_period = QPushButton("ТЕСТ Авторизация пользователей за период", self)
        self.authorization_in_period.move(0, 80)

        self.total_points = QPushButton("Общая информация по баллам на текущий момент", self)
        self.total_points.move(0, 115)
        self.total_points.clicked.connect(self.points_by_users_and_countries)

    @staticmethod
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
        total_stat_df.to_excel(f'total_stats_about_users_for_{datetime.now().date()}.xlsx')
        os.startfile(f'total_stats_about_users_for_{datetime.now().date()}.xlsx')

    @staticmethod
    def last_authorization_in_app():
        """ Information about last authorisation users in app by years"""

        df_users['Year'] = df_users['Последняя авторизация в приложении'].dt.year
        df_users['Year'].fillna(0, inplace=True)
        years = sorted([year for year in set(df_users['Year']) if year != 0])

        def amount_users_by_type(country_for_amount_users: str, user_type: str):
            """ Count amount users in country. """

            data = df_users[(df_users["Страна"] == country_for_amount_users) &
                            (df_users["Тип пользователя"] == user_type)]

            return len(data["ID"])

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

        columns = ['Страна', 'Тип пользователей', 'Всего в базе'] + [year for year in years] + ['Не авторизовались']
        index = [_ for _ in range(len(last_authorization_in_app_list))]
        last_authorization_in_app_df = pd.DataFrame(last_authorization_in_app_list, index, columns)

        last_authorization_in_app_df.to_excel(f'last_authorization_in_app {datetime.now().date()}.xlsx')
        os.startfile(f'last_authorization_in_app {datetime.now().date()}.xlsx')

    @staticmethod
    def points_by_users_and_countries():
        """ Information about points by users and countries """

        def sum_of_points(type_of_user: str, country_for_sum_points: str):
            """ Count point of users by country"""

            data = df_users[(df_users['Тип пользователя'] == type_of_user) &
                            (df_users['Страна'] == country_for_sum_points)]

            return sum(data['Баллы'])

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
