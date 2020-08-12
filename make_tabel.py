from PyQt5 import uic, QtWidgets
import tabel
import PyQt5
import pickle
import time
import sys
import os
import subprocess
# Объявление переменных
Form, _ = uic.loadUiType("tabel.ui")
topic = [0 for _ in range(5)]
topic_val = [0 for _ in range(5)]
rest = [0, 0]
holiday_list = []
komandirovka = []
ill = [0, 0]
users_list = {}
short_days_list = []
rest_word = {0: 'Отпуск', 1: 'Командировка', 2: 'Больничный'}
path_program = ''
path = ''

class Ui(QtWidgets.QMainWindow, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        global path, path_program
        self.update_bot.clicked.connect(self.update_it)
        self.make_tabel.clicked.connect(self.try_it)
        self.users.itemSelectionChanged.connect(self.selectionChanged)
        self.adduser.clicked.connect(self.add_user_click)
        self.read_users()
        self.removeuser.clicked.connect(self.remove_user_click)
        self.komand_add.clicked.connect(self.add_komandiroka)
        self.komand_remove.clicked.connect(self.remove_komandirovka)
        self.year.setText(str(time.localtime()[0]))
        if time.localtime()[2] > 9:
            self.month.setCurrentIndex(time.localtime()[1] - 1)
        else:
            self.month.setCurrentIndex(time.localtime()[1] - 2)
        path_program = os.getcwd()
        os.chdir('..')
        path = os.getcwd()

    # Обновление всех значений и перерасчёт дней
    def update_it(self):
        global topic, topic_val, user, year, month, rest, holiday_list, ill, komandirovka, short_days_list
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

        short_days_list = self.unworkdaysinput(self.short_days.toPlainText())
        holiday_list = self.unworkdaysinput(self.holiday.toPlainText())
        user = self.Name.toPlainText()
        year = int(self.year.toPlainText())
        month = self.month.currentIndex() + 1
        sum_top = 0
        sum_round = 0
        sum_month = tabel.count_hour(month, year, holiday_list, komandirovka, short_days_list)
        for i in range(5):
            if topic_val[i].isdigit():
                if int(topic_val[i]) + sum_round != 100:
                    sum_round += int(topic_val[i])
                    topic_val[i] = int(sum_month / 100 * int(topic_val[i]))
                    sum_top += topic_val[i]
                else:
                    topic_val[i] = sum_month - sum_top
                    sum_top += topic_val[i]
                    sum_round = 100
            else:
                topic_val[i] = 0
        self.sum_time.setText(f'{sum_round}')
        self.must_time.setText(f'Всего {sum_month} рабочих часов')
        self.waring_mes.setText('')

    # Создание табеля, сохранение пользователей в файл
    def try_it(self):
        self.update_it()
        os.chdir(path_program)
        try:
            tabel_name = tabel.make_tabel_func(user, month, year, topic, topic_val, holiday_list,
                                               komandirovka, short_days_list, path)
            self.waring_mes.setText('Табель успешно создан')
            if self.open_check.isChecked():
                os.startfile(tabel_name[0])
        except:
            self.waring_mes.setText('Закройте файл Excel')

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
                users_list_load = pickle.load(f)
            for i in users_list_load:
                if i not in users_list:
                    users_list_load[i].insert(0, PyQt5.QtWidgets.QListWidgetItem())
                    users_list_load[i][0].setText(i)
                    self.users.addItem(users_list_load[i][0])
                    users_list[i] = users_list_load[i]
        except:
            pass

    # Сохранение пользователей в файл
    def save_users(self, path):
        global users_list
        os.chdir(path)
        list_to_save = users_list.copy()
        for i in list_to_save:
            list_to_save[i].pop(0)
            self.users.takeItem(0)
        with open('users.tabel', 'wb') as f:
            pickle.dump(users_list, f)
        users_list = {}

    # Добавление пользователя
    def add_user_click(self):
        self.update_it()
        userinlist = False
        if user in users_list:
            userinlist = True
        users_list[user] = [PyQt5.QtWidgets.QListWidgetItem()] + topic
        users_list[user][0].setText(f'{user}')
        if not userinlist:
            self.users.addItem(users_list[user][0])
        self.save_users(path_program)
        self.read_users()

    # Удаление пользователя
    def remove_user_click(self):
        try:
            users_list.pop(self.users.takeItem(self.users.currentRow()).text())
        except:
            pass

    # Добавление строки отпуска, командировки или больничного
    def add_komandiroka(self):
        self.update_it()
        komandirovka_days = []
        max_days = len([i for i in tabel.workday(1, month, year, [], [])]) + 1
        for day in komandirovka:
            komandirovka_days += [i for i in range(day[1], day[2] + 1)]

        if self.komand_start.text().isdigit() and self.komand_end.text().isdigit():
            if max_days > int(self.komand_end.text()) >= int(self.komand_start.text()):
                local_days = [i for i in range(int(self.komand_start.text()), int(self.komand_end.text()) + 1)]
                days_in_komandirovka = False
                for i in local_days:
                    if i in komandirovka_days:
                        days_in_komandirovka = True
                if not days_in_komandirovka:
                    k = tuple([self.komand_comboBox.currentIndex(), int(self.komand_start.text()),
                               int(self.komand_end.text())])
                    komandirovka.append(k)
                    self.komand_list.addItem(f'{rest_word[k[0]]} с {k[1]} по {k[2]} число')
                    komandirovka_days += local_days
        self.komand_start.setText('')
        self.komand_end.setText('')
        self.update_it()

    # Удаление строки отпуска, командировки или больничного
    def remove_komandirovka(self):
        if self.komand_list.currentRow() != -1:
            try:
                komandirovka.pop(self.komand_list.currentRow())
                self.komand_list.takeItem(self.komand_list.currentRow())
                self.update_it()
            except:
                pass
            self.update_it()

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

    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
