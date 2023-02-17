from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from functions import *


class MainWindow(QDialog):
    df_users = None
    countries = None

    def __init__(self):
        super().__init__()
        # set Title for main windows
        self.setWindowTitle("Данные по пользователя и сканам в приложении AXOR")
        self.setWindowIcon(QIcon('axor.ico'))
        # # Set size of main window
        self.resize(600, 400)

        self.about_users_btn = QPushButton("Загрузить базу пользователей", self)
        self.about_users_btn.move(0, 20)
        self.about_users_btn.clicked.connect(self.check_file_with_users)

        self.users_in_countries_bt = QPushButton("Пользователи по странам", self)
        self.users_in_countries_bt.move(0, 50)
        self.users_in_countries_bt.clicked.connect(self.users_by_country)

        self.authorization_in_app_btn = QPushButton("Авторизация пользователей в приложении", self)
        self.authorization_in_app_btn.move(0, 85)
        self.authorization_in_app_btn.clicked.connect(self.last_authorization_in_app)

        self.authorization_in_period_btn = QPushButton("ТЕСТ Авторизация пользователей за период", self)
        self.authorization_in_period_btn.move(0, 120)

        self.total_points_btn = QPushButton("Общая информация по баллам на текущий момент", self)
        self.total_points_btn.move(0, 155)
        self.total_points_btn.clicked.connect(self.points_by_users_and_countries)

        self.total_points_btn = QPushButton("Кол-во сканировавших пользователей в текущем году на данный момент", self)
        self.total_points_btn.move(0, 190)
        self.total_points_btn.clicked.connect(self.data_about_scan_users_in_current_year)

    def check_file_with_users(self):
        """Loading and check file about users and the availability necessary columns in file about users"""
        global df_users
        global countries

        file_with_data = QFileDialog.getOpenFileName(self, 'Open file', 'C:/', '*.*')

        print("Загрузка данных по пользователям...")

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

        data_about_users = pd.read_excel(file_with_data[0],
                                         na_values="NA",
                                         converters={"ID": int, "Баллы": int})

        for col_name in df_users_columns:
            if col_name not in data_about_users.columns:
                QMessageBox.warning(self, "Внимание!", f"В загруженных данных не хватает столбца {col_name}")
                return False
            else:
                data_about_users = data_about_users[['ID',
                                                     'Баллы',
                                                     'Последняя авторизация в приложении',
                                                     'Город работы',
                                                     'Страна',
                                                     'Тип пользователя',
                                                     'Фамилия',
                                                     'Имя',
                                                     'Отчество',
                                                     'E-Mail']]

        data_about_users['Последняя авторизация в приложении'] = pd.to_datetime(
                                                                data_about_users['Последняя авторизация в приложении'],
                                                                format='%d.%m.%Y %H:%M:%S').dt.normalize()

        data_about_users['Баллы'].fillna(0, inplace=True)
        data_about_users.fillna('', inplace=True)

        """Clean spam (exception empty row in "Страна" and 'Клиент' as spam) and test accounts in DataFrame"""
        data_about_users = data_about_users[(data_about_users['Страна'] != '') &
                                            (data_about_users['Тип пользователя'] != 'Клиент')]

        """List of test accounts, excludes from counting"""
        exclude_users = ['kazah89', 'sanin, ''samoilov', 'axorindustry', 'kreknina', 'zeykin', 'berdnikova',
                         'ostashenko',
                         'skalar', 'test', 'malyigor', 'ihormaly', 'axor',
                         'kosits']

        """Creating list of excluded accounts"""
        exclude_list = set()
        for email in data_about_users['E-Mail']:
            for i in exclude_users:
                if i in email:
                    exclude_list.add(email)

        """Clean DataFrame from exclude accounts"""
        data_about_users = data_about_users.loc[~data_about_users['E-Mail'].isin(exclude_list)]
        df_users = data_about_users
        countries = list(set(df_users["Страна"]))  # list of countries in DataFrame

        QMessageBox.information(self, "Информация", "Данные по пользователям загружены.")

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

    @staticmethod
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
