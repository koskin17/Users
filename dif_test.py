from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QVBoxLayout, QWidget, QPushButton

import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("My App")

        self.btn_about_users = QPushButton("Загрузить базу пользователей", self)
        self.btn_about_users.move(0, 175)

        self.btn_users_by_country = QPushButton("Пользователи по странам", self)
        self.btn_users_by_country.move(0, 205)

        layout = QVBoxLayout()
        layout.addWidget(self.btn_about_users)
        layout.addWidget(self.btn_users_by_country)

        container = QWidget()
        container.setLayout(layout)

        # Устанавливаем центральный виджет Window.
        self.setCentralWidget(container)


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()