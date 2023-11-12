from PyQt6.QtWidgets import QWidget, QApplication, QLineEdit, QPushButton, QListWidget, QPlainTextEdit, QLabel
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QRect, Qt
import sys
import pyexcel


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedWidth(600)
        self.setFixedHeight(350)
        self.setWindowIcon(QIcon('icons/icon.jpg'))
        self.setWindowTitle('Проект Алексея Полочкина "Потеряшка"')

        self.find_text = QLineEdit(self)
        self.find_text.setGeometry(QRect(10, 30, 491, 20))
        self.find_text.setPlaceholderText('Фамилия Имя')
        self.find_text.returnPressed.connect(lambda: self.find_pupil())
        self.sendButton = QPushButton('Найти', self)
        self.sendButton.setGeometry(QRect(510, 10, 75, 23))
        self.sendButton.clicked.connect(lambda: self.find_pupil())
        self.sendButton.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dark_theme = QPushButton('', self)
        self.dark_theme.setGeometry(QRect(550, 190, 41, 41))
        self.dark_theme.clicked.connect(lambda: self.dark())
        self.dark_theme.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dark_theme.setIcon(QIcon('icons/dark_theme.jpg'))
        self.label = QLabel('Кого ищем?', self)
        self.label.setGeometry(QRect(12, 10, 480, 16))
        self.label_2 = QLabel('Найденные ученики: ', self)
        self.label_2.setGeometry(QRect(12, 60, 491, 16))
        self.listView = QListWidget(self)
        self.listView.setGeometry(QRect(10, 80, 491, 241))
        self.conclusion = QPlainTextEdit('Project was made by young genius.', self)
        self.conclusion.setGeometry(QRect(505, 250, 91, 81))
        self.conclusion.setReadOnly(True)
        self.setStyleSheet('background-color: #f0f0f0;')
        self.conclusion.setStyleSheet(
            "QPlainTextEdit {background-color: #f0f0f0; border: none; text-align: center; color: #000;}")
        self.find_text.setStyleSheet("QLineEdit{color: #000; border-radius: 5px; background-color: #fafafa;}")
        self.sendButton.setStyleSheet("QPushButton{background-color: #fff; border-radius: 5px;}")
        self.listView.setStyleSheet(
            "QListWidget{background-color: #fff; color: #000; border-radius: 5px; border: 0px solid #000;}")
        self.dark_theme.setStyleSheet("QPushButton{background-color: #fff; border-radius: 5px;}")

        self.flag = True

    def dark(self):
        if self.flag:
            self.setStyleSheet('background-color: #18181b;')
            self.conclusion.setStyleSheet(
                "QPlainTextEdit {background-color: #18181b; border: none; text-align: center; color: #fff;}")
            self.find_text.setStyleSheet("QLineEdit{color: #fff; border-radius: 5px;}")
            self.sendButton.setStyleSheet("QPushButton{background-color: #f0f0f0; border-radius: 5px;}")
            self.listView.setStyleSheet("QListWidget{background-color: #393939; color: #f0f0f0; border-radius: 5px;}")
            self.dark_theme.setStyleSheet("QPushButton{background-color: #fff; border-radius: 5px;}")
            self.dark_theme.setIcon(QIcon('icons/light_theme.jpg'))
            self.label.setStyleSheet("QLabel{color: #fff;}")
            self.label_2.setStyleSheet("QLabel{color: #fff;}")
        else:
            self.setStyleSheet('background-color: #f0f0f0;')
            self.conclusion.setStyleSheet(
                "QPlainTextEdit {background-color: #f0f0f0; border: none; text-align: center; color: #000;}")
            self.find_text.setStyleSheet("QLineEdit{color: #000; border-radius: 5px;}")
            self.sendButton.setStyleSheet("QPushButton{background-color: #fff; border-radius: 5px;}")
            self.listView.setStyleSheet(
                "QListWidget{background-color: #fff; color: #000; border-radius: 5px; border: 0px solid #000;}")
            self.dark_theme.setStyleSheet("QPushButton{background-color: #fff; border-radius: 5px;}")
            self.dark_theme.setIcon(QIcon('icons/dark_theme.jpg'))
            self.label.setStyleSheet("QLabel{color: #000;}")
            self.label_2.setStyleSheet("QLabel{color: #000;}")
        self.flag = not self.flag

    def find_pupil(self):
        counter = 0
        self.listView.clear()
        text = self.find_text.text()
        if text == '':
            return
        file_name = 'pupils.xls'
        pupils = pyexcel.get_array(file_name=file_name)

        for pupil in pupils:
            if text.lower() in pupil[0].lower():
                self.listView.addItem(f'{pupil[0]}: класс {pupil[1]}, классная комната {pupil[2]}')
                counter += 1
        if counter == 0:
            self.listView.addItem('Никого не найдено...')


if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
