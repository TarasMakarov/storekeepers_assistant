import sys

import pandas as pd

from PyQt6.QtWidgets import QMainWindow, QApplication, QPushButton, QLineEdit, QFileDialog, QLabel, QCheckBox


class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowTitle('Помощник кладовщика')
        self.setGeometry(400, 300, 700, 500)

        self.main_text = QLabel(self)
        self.main_text.setText('Провести следующие операции:')
        self.main_text.move(50, 70)
        self.main_text.adjustSize()

        self.btn_choose_file = QPushButton(self)
        self.btn_choose_file.setGeometry(550, 25, 80, 25)
        self.btn_choose_file.setText('Выбрать файл')
        self.btn_choose_file.adjustSize()
        self.btn_choose_file.clicked.connect(self.search_files)

        self.line = QLineEdit(self)
        self.line.setGeometry(50, 25, 485, 25)

        self.btn_check_find_vc_in_two_places = QCheckBox(self)
        self.btn_check_find_vc_in_two_places.setChecked(False)
        self.btn_check_find_vc_in_two_places.setGeometry(50, 120, 15, 15)

        self.label_check_find_vc_in_two_places = QLabel(self)
        self.label_check_find_vc_in_two_places.setText('Найти товар, размещенный в двух и более ячейках')
        self.label_check_find_vc_in_two_places.move(70, 120)
        self.label_check_find_vc_in_two_places.adjustSize()

        self.btn_check_two_or_more_vc_in_one_place = QCheckBox(self)
        self.btn_check_two_or_more_vc_in_one_place.setChecked(False)
        self.btn_check_two_or_more_vc_in_one_place.setGeometry(50, 150, 15, 15)

        self.label_check_two_or_more_vc_in_one_place = QLabel(self)
        self.label_check_two_or_more_vc_in_one_place.setText('Найти ячейки, в которых больше одного вида товара (НЕ ГОТОВО)')
        self.label_check_two_or_more_vc_in_one_place.move(70, 150)
        self.label_check_two_or_more_vc_in_one_place.adjustSize()

        self.btn_check_empty_bottom_cells = QCheckBox(self)
        self.btn_check_empty_bottom_cells.setChecked(False)
        self.btn_check_empty_bottom_cells.setGeometry(50, 180, 15, 15)

        self.label_check_empty_bottom_cells = QLabel(self)
        self.label_check_empty_bottom_cells.setText('Найти пустые нижние ячейки на складах отгрузки (НЕ ГОТОВО)')
        self.label_check_empty_bottom_cells.move(70, 180)
        self.label_check_empty_bottom_cells.adjustSize()

        self.btn_run_action = QPushButton(self)
        self.btn_run_action.setGeometry(50, 420, 80, 25)
        self.btn_run_action.setText('Начать')
        self.btn_run_action.clicked.connect(self.run_actions)

        self.file = ''
        self.rows = -3
        self.f = pd.DataFrame()

        self.duplicates = dict()  # for def find_vc_in_two_places
        self.duplicates_res = dict()  # for def find_vc_in_two_places

        self.result = pd.DataFrame()

    def search_files(self):
        f_name = QFileDialog.getOpenFileName(self)
        self.file = f_name[0]
        self.open_file()

    def open_file(self):
        self.f = pd.read_excel(self.file)
        self.rows += self.f.shape[0]
        self.line.setText(self.file)

    def find_vc_in_two_places(self):
        row = 2
        for i in range(self.rows):
            temp_list = []
            s = self.f.iloc[row, 1]
            if isinstance(s, str):
                temp_list = s.split(', ')
            artikul = str(self.f.iloc[row, 3]).split('.')[0]
            if artikul not in self.duplicates.keys():
                self.duplicates[artikul] = temp_list
                row += 1
            else:
                for t in temp_list:
                    if t not in self.duplicates[artikul]:
                        self.duplicates[artikul].append(t)
                row += 1
        for k, v in self.duplicates.items():
            if len(v) > 1:
                self.duplicates_res[k] = v
        return self.duplicates_res

    def find_two_or_more_vc_in_one_place(self):
        pass

    def find_empty_bottom_cells(self):
        pass

    def run_actions(self):
        if self.btn_check_find_vc_in_two_places.isChecked():
            self.find_vc_in_two_places()
        if self.btn_check_two_or_more_vc_in_one_place.isChecked():
            self.find_two_or_more_vc_in_one_place()
        top = pd.DataFrame({'одинаковый товар в нескольких ячейках'})
        self.result = pd.DataFrame({'код товара': self.duplicates_res.keys(), 'ячейки': self.duplicates_res.values()})
        with pd.ExcelWriter('/home/taras_v_m/PycharmProjects/storekeeper\'s_assistant/result.xlsx') as writer:
            top.to_excel(writer)
            self.result.to_excel(writer, startrow=2)
        self.window().close()

    def add_action(self):
        self.btn_check_find_vc_in_two_places.clicked.connect(lambda: self.find_vc_in_two_places)
        self.btn_check_two_or_more_vc_in_one_place.clicked.connect(lambda: self.find_two_or_more_vc_in_one_place)
        self.btn_run_actions.clicked.connect(self.run_actions)


def application():
    app = QApplication(sys.argv)
    window = Window()

    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    application()