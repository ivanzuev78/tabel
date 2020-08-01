from PyQt5 import uic, QtWidgets
# from PyQt5.QtWidgets import QApplication
import tabel


Form, _ = uic.loadUiType("tabel.ui")
topic = [0 for _ in range(4)]
topic_val = [0 for _ in range(4)]
user = None
year = 2020
month = 1


class Ui(QtWidgets.QMainWindow, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        # self.add_user.clicked.connect(self.add_user_func)
        self.update_bot.clicked.connect(self.update_it)
        self.make_tabel.clicked.connect(self.try_it)

    def update_it(self):
        global topic, topic_val, user, year, month
        topic[0] = self.top_1.toPlainText()
        topic[1] = self.top_2.toPlainText()
        topic[2] = self.top_3.toPlainText()
        topic[3] = self.top_4.toPlainText()
        topic_val[0] = self.top_1_val.toPlainText()
        topic_val[1] = self.top_2_val.toPlainText()
        topic_val[2] = self.top_3_val.toPlainText()
        topic_val[3] = self.top_4_val.toPlainText()
        user = self.Name.toPlainText()
        year = int(self.year.toPlainText())
        month = self.month.currentIndex() + 1
        sum_top = 0
        for i in range(4):
            if topic_val[i].isdigit():
                sum_top += int(topic_val[i])
        self.sum_time.setText(f'{sum_top}')
        print(month, year, tabel.count_hour(month, year))
        self.must_time.setText(f'{tabel.count_hour(month, year)}')

    def try_it(self):
        tabel.make_tabel_func(user, month, year, topic, topic_val, holiday=[])
        sys.exit()


if __name__ == "__main__":

    import sys
    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()  # show window
    sys.exit(app.exec_())
