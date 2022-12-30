from functions import *


def main():
    choice = ''
    while choice != 'q':
        print()
        print('Информация по пользователям:')
        print()
        print('1: Пользователи по странам')
        print('2: Авторизация пользователей')
        print('3: Авторизация пользователей за период')
        print('4: Общая информация по баллам на текущий момент')
        print()
        print('Информация по сканам:')
        print()
        print('7. Кол-во сканировавших пользователей в текущем году на данный момент.')
        print('8. Данные по насканированным баллам в текущем году на данный момент.')
        print('9. ТОП дилеров по сканам в текущем году на данный момент.')
        print('10. ТОП монтажников по сканам в текущем году на данный момент.')
        print('11. Кол-во пользователей и насканированных баллов за период.')
        print('12. Кол-во сканировавших пользователей в текущем году по месяцам.')
        print()
        print('Q: Завершить работу')
        print()

        choice = input('Укажите номер пункта: ').lower()
        match choice:
            case '1':
                total_stat()
            case '2':
                last_authorization_in_app()
            case '3':
                start_date = input('Укажите дату начала периода в формате mm.dd.yyyy (через точку): ')
                try:
                    valid_start_date = time.strptime(start_date, '%d.%m.%Y')
                    start_date = time.strftime('%d.%m.%Y', valid_start_date)
                    end_date = input('Укажите дату конца периода в формате mm.dd.yyyy (через точку): ')
                    try:
                        valid_end_date = time.strptime(end_date, '%d.%m.%Y')
                        end_date = time.strftime('%d.%m.%Y', valid_end_date)
                        authorization_during_period(start_date, end_date)
                    except ValueError:
                        print('Дата введена неверно!')
                except ValueError:
                    print('Дата введена неверно!')
            case '4':
                points_by_users_and_countries()
            case '7':
                table_about_scan_users_in_year()
            case '8':
                data_about_points()
            case '9':
                country_list = {}
                for i in range(len(countries)):
                    print(i, ' - ', countries[i])
                    country_list[i] = countries[i]
                    i += 1

                country_choice = input('Выберите страну: ')
                top_users_by_scans(country_list[int(country_choice)], 'Дилер')
            case '10':
                country_list = {}
                for i in range(len(countries)):
                    print(i, ' - ', countries[i])
                    country_list[i] = countries[i]
                    i += 1

                country_choice = input('Выберите страну: ')
                top_users_by_scans(country_list[int(country_choice)], 'Монтажник')
            case '11':
                start_date = input('Укажите дату начала периода в формате mm.dd.yyyy (через точку): ')
                try:
                    valid_start_date = time.strptime(start_date, '%d.%m.%Y')
                    start_date = time.strftime('%d.%m.%Y', valid_start_date)
                    end_date = input('Укажите дату конца периода в формате mm.dd.yyyy (через точку): ')
                    try:
                        valid_end_date = time.strptime(end_date, '%d.%m.%Y')
                        end_date = time.strftime('%d.%m.%Y', valid_end_date)
                        data_about_scans_during_period(start_date, end_date)
                    except ValueError:
                        print('Дата введена неверно!')
                except ValueError:
                    print('Дата введена неверно!')
            case '12':
                table_scanned_users_by_months()
            case 'q':
                finish()


''' НАЧАЛО ОСНОВНОЙ ПРОГРАММЫ '''

main()
