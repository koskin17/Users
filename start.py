from classes import *
# from functions import *
import sys


def main():
    # Create application app.
    app = QApplication(sys.argv)
    # Create main window for all other widgets.
    main_window = MainWindows()
    # Showing the main windows
    main_window.show()
    # Start app with method exec_, which starting loop
    sys.exit(app.exec_())

    choice = ''
    while choice != 'q':
        print()
        print('Информация по пользователям:')
        print()
        print('3: Авторизация пользователей за период')
        print()
        print('Информация по сканам:')
        print()
        print('5. Кол-во сканировавших пользователей в текущем году на данный момент.')
        print('6. Данные по насканированным баллам в текущем году на данный момент.')
        print('7. ТОП дилеров по сканам в текущем году на данный момент.')
        print('8. ТОП монтажников по сканам в текущем году на данный момент.')
        print('9. Кол-во пользователей и насканированных баллов за период.')
        print('10. Кол-во сканировавших пользователей в текущем году по месяцам.')
        print()
        print('Q: Завершить работу')
        print()

        choice = input('Укажите номер пункта: ').lower()
        match choice:
            case '3':
                try:
                    start_date = datetime.strptime(
                        input('Укажите дату начала периода в формате mm.dd.yyyy (через точку): '), '%d.%m.%Y')
                    try:
                        end_date = datetime.strptime(
                            input('Укажите дату конца периода в формате mm.dd.yyyy (через точку): '), '%d.%m.%Y')
                        authorization_during_period(start_date, end_date)
                    except ValueError:
                        print('Конечная дата периода введена неверно!')
                except ValueError:
                    print('Начальная дата периода введена неверно!')
            case '5':
                data_about_scan_users_in_current_year()
            case '6':
                data_about_points()
            case '7':
                country_list = {}
                for _ in range(len(countries)):
                    print(_, ' - ', countries[_])
                    country_list[_] = countries[_]
                    _ += 1

                country_choice = input('Выберите страну: ')
                top_users_by_scans(country_list[int(country_choice)], 'Дилер')
            case '8':
                country_list = {}
                for _ in range(len(countries)):
                    print(_, ' - ', countries[_])
                    country_list[_] = countries[_]
                    _ += 1

                country_choice = input('Выберите страну: ')
                top_users_by_scans(country_list[int(country_choice)], 'Монтажник')
            case '9':
                try:
                    start_date = datetime.strptime(
                        input('Укажите дату начала периода в формате mm.dd.yyyy (через точку): '), '%d.%m.%Y')
                    try:
                        end_date = datetime.strptime(
                            input('Укажите дату конца периода в формате mm.dd.yyyy (через точку): '), '%d.%m.%Y')
                        data_about_scans_during_period(start_date, end_date)
                    except ValueError:
                        print('Дата введена неверно!')
                except ValueError:
                    print('Дата введена неверно!')
            case '10':
                scanned_users_by_months()
            case 'q':
                finish()


''' НАЧАЛО ОСНОВНОЙ ПРОГРАММЫ '''
if __name__ == '__main__':
    if check_file_with_users() and check_file_with_scans():
        main()
