from PyQt5 import uic, QtWidgets
import tabel
import PyQt5
import pickle
import time


# Объявление переменных
Form, _ = uic.loadUiType("tabel.ui")
topic = [0 for _ in range(5)]
topic_val = [0 for _ in range(5)]
rest = [0, 0]
holiday_list = []
komandirovka = []
ill = [0, 0]
users_list = {}


class Ui(QtWidgets.QMainWindow, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        # self.add_user.clicked.connect(self.add_user_func)
        self.update_bot.clicked.connect(self.update_it)
        self.make_tabel.clicked.connect(self.try_it)
        self.users.itemSelectionChanged.connect(self.selectionChanged)
        self.adduser.clicked.connect(self.add_user_click)
        self.read_users()
        self.removeuser.clicked.connect(self.remove_user_click)
        self.year.setText(str(time.localtime()[0]))
        self.month.setCurrentIndex(time.localtime()[1] - 1)
    #  Выбор пользователя в таблице
    def selectionChanged(self):
        try:
            self.top_1.setText(users_list[self.users.selectedItems()[0].text()][1])
            self.top_2.setText(users_list[self.users.selectedItems()[0].text()][2])
            self.top_3.setText(users_list[self.users.selectedItems()[0].text()][3])
            self.top_4.setText(users_list[self.users.selectedItems()[0].text()][4])
            self.top_5.setText(users_list[self.users.selectedItems()[0].text()][5])
            self.Name.setText(self.users.selectedItems()[0].text())
        except:
            pass

    # Чтение пользователей из файла
    def read_users(self):
        global users_list
        try:
            with open('users.tabel', 'rb') as f:
                users_list = pickle.load(f)
            for i in users_list:
                users_list[i].insert(0, PyQt5.QtWidgets.QListWidgetItem())
                users_list[i][0].setText(i)
                self.users.addItem(users_list[i][0])
        except:
            pass

    # Сохранение пользователей в файл
    def save_users(self):
        for i in users_list:
            users_list[i].pop(0)
        with open('users.tabel', 'wb') as f:
            pickle.dump(users_list, f)

    # Обновление всех значений и перерасчёт дней
    def update_it(self):
        global topic, topic_val, user, year, month, rest, holiday_list
        topic[0] = self.top_1.toPlainText()
        topic[1] = self.top_2.toPlainText()
        topic[2] = self.top_3.toPlainText()
        topic[3] = self.top_4.toPlainText()
        topic[4] = self.top_5.toPlainText()
        topic_val[0] = self.top_1_val.toPlainText()
        topic_val[1] = self.top_2_val.toPlainText()
        topic_val[2] = self.top_3_val.toPlainText()
        topic_val[3] = self.top_4_val.toPlainText()
        topic_val[4] = self.top_5_val.toPlainText()
        if self.restcheckBox.isChecked():
            try:
                rest[0] = int(self.rest_start.text())
                rest[1] = int(self.rest_end.text())
            except:
                pass

        if self.boln_checkBox.isChecked():
            try:
                ill[0] = int(self.boln_start.text())
                ill[1] = int(self.boln_end.text())
            except:
                pass

        user = self.Name.toPlainText()
        year = int(self.year.toPlainText())
        month = self.month.currentIndex() + 1
        sum_top = 0
        for i in range(5):
            if topic_val[i].isdigit():
                topic_val[i] = int(topic_val[i])
                sum_top += topic_val[i]
            else:
                topic_val[i] = 0
        holiday_list = self.unworkdaysinput(self.holiday.toPlainText())
        self.sum_time.setText(f'{sum_top}')
        self.must_time.setText(f'{tabel.count_hour(month, year, rest, holiday_list)}')

    # Добавление пользователя
    def add_user_click(self):
        self.update_it()
        users_list[user] = [PyQt5.QtWidgets.QListWidgetItem()] + topic
        users_list[user][0].setText(f'{user}')
        self.users.addItem(users_list[user][0])

    # Удаление пользователя
    def remove_user_click(self):
        # print(self.users.takeItem(self.users.currentRow()).text())
        try:
            users_list.pop(self.users.takeItem(self.users.currentRow()).text())
        except:
            pass

    # Создание табеля, сохранение пользователей в файл
    def try_it(self):
        self.update_it()
        self.save_users()
        tabel.make_tabel_func(user, month, year, topic, topic_val, holiday_list, rest)
        sys.exit()

    # Распознавание текста в сокращенных и нерабочих днях
    def unworkdaysinput(self, inp):
        l = len(inp)
        integ = []
        i = 0
        while i < l:
            s_int = ''
            a = inp[i]
            while '0' <= a <= '9':
                s_int += a
                i += 1
                if i < l:
                    a = inp[i]
                else:
                    break
            i += 1
            if s_int != '':
                integ.append(int(s_int))
        return integ

if __name__ == "__main__":

    import sys
    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
