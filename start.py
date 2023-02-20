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
        print()


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
