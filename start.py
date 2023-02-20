from classes import *
import sys


def main():
    choice = ''
    while choice != 'q':
        print()
        print()
        print()
        print()
        print()
        print()
        print()
        print()
        print('9. Кол-во пользователей и насканированных баллов за период.')

        choice = input('Укажите номер пункта: ').lower()
        match choice:
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


''' Start main program '''
if __name__ == '__main__':
    # Create application app.
    app = QApplication(sys.argv)
    # Create main window for all other widgets.
    main_window = MainWindow()
    main_window.windowIcon()
    # Showing the main windows
    main_window.show()
    # Start app with method exec_, which starting loop
    sys.exit(app.exec_())
    # if not df_users.empty and not df_scans.empty:
    #     main()
