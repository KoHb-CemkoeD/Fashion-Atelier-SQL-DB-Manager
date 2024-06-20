import sys
from builtins import enumerate
from datetime import datetime
from os import system, curdir
from os.path import abspath

from PyQt5.QtCore import QDateTime, QModelIndex
from PyQt5 import uic, Qt, QtCore
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QMainWindow, QApplication, QAction, QFileDialog, QHeaderView, QCompleter

from res.main_classes import DBManager, Report

INSERT_TYPE, DELETE_TYPE, SELECT_TYPE, UPDATE_TYPE = 0, 1, 2, 3


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi("res/main_layout.ui", self)
        self.fields = {}
        self.search = False
        self.last_report_path = ""
        self.type_act.addItems(["Добавление записей", "Удаление данных", "Выборка данных", "Изменение данных"])
        self.combo_types, self.tables, self.field_types, self.radio_butts, self.db_manager, self.completer = \
            [], [], [], [], DBManager(), QCompleter([], self)
        self.lines_cur, self.lines_cond, self.lines_mod = [], [], []
        self.completer.setCaseSensitivity(Qt.Qt.CaseInsensitive)
        self.lineEdit_1.setCompleter(self.completer)
        self.db_manager.transform_data = self.transform_data
        self.types_spl_sizes = [[90, 43, 0, 0, 45], [135, 43, 43, 0, 95], [175, 43, 43, 43, 135], [90, 43, 0, 0, 45]]
        self.fields_query = [["артикул_модели", "назначение", "спецификация", "размер"],
                             ["заказчик", "заказы", "название", "размер"],
                             ["Код", "ФИО", "телефон", "адрес", "специализация", "стаж_работы"],
                             ["заказчик", "номер_заказа", "дата", "закройщик", "заказ", "размер"],
                             ["заказчик", "код", "назначение", "размер", "дата_готовности", "тип"]]

        self.queries = ["SELECT * FROM МОДЕЛИ",

                        "SELECT заказчик, дата, название, размер FROM ЗАКАЗЫ ORDER BY заказчик, дата",

                        "SELECT * FROM ЗАКРОЙЩИКИ",

                        """SELECT ЗАКАЗЫ.заказчик, ЗАКАЗЫ.номер, ЗАКАЗЫ.дата, закройщики.ФИО, ЗАКАЗЫ.название, ЗАКАЗЫ.размер
FROM ЗАКАЗЫ INNER JOIN ЗАКРОЙЩИКИ ON ЗАКАЗЫ.закройщик = ЗАКРОЙЩИКИ.Код 
{}
ORDER BY ЗАКАЗЫ.заказчик, ЗАКАЗЫ.дата""",

                        """SELECT ГОТОВАЯ_ПРОДУКЦИЯ.заказчик, ГОТОВАЯ_ПРОДУКЦИЯ.код, МОДЕЛИ.назначение, ГОТОВАЯ_ПРОДУКЦИЯ.размер, ГОТОВАЯ_ПРОДУКЦИЯ.дата_готовности, ТКАНИ.тип
FROM ((ГОТОВАЯ_ПРОДУКЦИЯ INNER JOIN МОДЕЛИ ON ГОТОВАЯ_ПРОДУКЦИЯ.модель=МОДЕЛИ.артикул_модели) 
INNER JOIN ТКАНИ ON ГОТОВАЯ_ПРОДУКЦИЯ.материал=ТКАНИ.артикул) 
{}
ORDER BY ГОТОВАЯ_ПРОДУКЦИЯ.заказчик, ГОТОВАЯ_ПРОДУКЦИЯ.код"""]

        self.postfixes = ["Всего моделей: ", "Всего заказчиков: ", "Всего закройщиков: ",
                          "Всего заказов: ", "Всего готовой продукции: "]
        self.init_controls()
        self.show()

    def tab_window_changed(self):
        cur_window = self.tabWidget.currentIndex()
        if cur_window == 1:
            self.resize_to_type(0, 0, 0, 0, 0)
        elif cur_window == 2:
            self.set_data_table(self.statictics_table, None,
                                ["Запрос", "Дата", "Результат"], self.db_manager.statistics, stretch=[0])
        else:
            try:
                checked_but = [i for i in range(5) if self.radio_butts[i].isChecked()][0]
                model = self.completer.model()
                model.setStringList(list(set(self.db_manager.exec_and_get_data("SELECT заказчик FROM ЗАКАЗЫ"))))
                self.radio_pos_changed(checked_but, True)
            except Exception as e:
                print(e)

    def resize_to_type(self, box, f_sp, c_sp, m_sp, sp_6):
        self.groupBox.setMinimumHeight(box)
        self.groupBox.setMaximumHeight(box)
        self.fields_sp.setMinimumHeight(f_sp)
        self.fields_sp.setMaximumHeight(f_sp)
        self.cond_sp.setMinimumHeight(c_sp)
        self.cond_sp.setMaximumHeight(c_sp)
        self.mod_sp.setMinimumHeight(m_sp)
        self.mod_sp.setMaximumHeight(m_sp)
        self.splitter_6.setMaximumHeight(sp_6)
        self.splitter_6.setMinimumHeight(sp_6)

    def hide_fields(self, cur_fields, count_fields, label, field, splitter):
        for i, name in cur_fields:
            eval(label + str(i + 1) + ".setText('" + name + "')")
            eval(field + str(i + 1) + ".clear()")
            eval(splitter + str(i + 1) + ".setMaximumWidth(999)")
        for i in range(count_fields, 9):
            eval(splitter + str(i) + ".setMaximumWidth(0)")

    def query_type_pos_changed(self):
        try:
            self.statusBar().showMessage("")
            cur_type = self.type_act.currentIndex()
            cur_table = self.table_act.currentText()
            cur_tab_index = self.table_act.currentIndex()
            cur_fields = list(enumerate(self.fields[cur_table]))
            cur_model = self.db_manager.get_model()
            count_fields = len(self.fields[cur_table]) + 1
            if self.create_but.text() == "Сбросить результаты":
                self.create_but.setText("Выполнить запрос")
                self.table_pos_changed()
            if cur_type == 0:
                self.resize_to_type(90, 43, 0, 0, 45)
                self.hide_fields(cur_fields, count_fields, "self.label_field_", "self.field_", "self.splitter_f")
            elif cur_type == 1 or cur_type == 2:
                self.resize_to_type(135, 43, 43, 0, 95)
                self.hide_fields(cur_fields, count_fields, "self.label_field_", "self.field_", "self.splitter_f")
                self.hide_fields(cur_fields, count_fields, "self.label_cond_", "self.cond_", "self.splitter_c")
            elif cur_type == 3:
                self.resize_to_type(175, 43, 43, 43, 135)
                self.hide_fields(cur_fields, count_fields, "self.label_field_", "self.field_", "self.splitter_f")
                self.hide_fields(cur_fields, count_fields, "self.label_cond_", "self.cond_", "self.splitter_c")
                self.hide_fields(cur_fields, count_fields, "self.label_mod_", "self.mod_", "self.splitter_m")
            for ind_field, field in cur_fields:
                cur_model.setHeaderData(ind_field, Qt.Qt.Horizontal, field)
                if self.field_types[cur_tab_index][ind_field] == "CHAR":
                    self.lines_cond[ind_field].setPlaceholderText("Условие(=)")
                else:
                    self.lines_cond[ind_field].setPlaceholderText("Условие(=, >, <, >=, <=)")
                if self.field_types[cur_tab_index][ind_field] == "DATE":
                    self.lines_cur[ind_field].setPlaceholderText("Текущ. значение(DD.MM.YYYY)")
                    self.lines_mod[ind_field].setPlaceholderText("Новое значение(DD.MM.YYYY)")
                else:
                    self.lines_cur[ind_field].setPlaceholderText("Текущ. значение")
                    self.lines_mod[ind_field].setPlaceholderText("Новое значение")
        except Exception as e:
            print(e)

    def table_pos_changed(self):
        try:
            if self.create_but.text() == "Сбросить результаты":
                self.create_but.setText("Выполнить запрос")
            self.table.reset()
            cur_tab_index = self.table_act.currentIndex()
            model = self.db_manager.get_model_by_index(cur_tab_index)
            self.table.setModel(model)
            for i in range(model.columnCount()):
                if self.field_types[cur_tab_index][i] == "CHAR":
                    self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            self.query_type_pos_changed()
        except Exception as e:
            print(e)

    def radio_pos_changed(self, index, restore=False):
        self.statusBar().showMessage("")
        if not restore:
            self.lineEdit_1.clear()
            self.lineEdit_2.clear()
            self.lineEdit_3.clear()
            self.table_query_res.clearSpans()
            self.table_query_res.reset()
            self.textEdit.setText(self.queries[index].format(""))
            self.lineEdit_2.setPlaceholderText("")
            self.lineEdit_3.setPlaceholderText("")
            self.table_query_res.setModel(QStandardItemModel(self))
        if index < 3:
            self.splitter_5.setMaximumWidth(0)
            self.splitter_8.setMaximumWidth(0)
            self.splitter_9.setMaximumWidth(0)
        elif index == 3:
            self.label_5.setText("Заказчик")
            self.label_6.setText("Начало периода")
            self.label_7.setText("Конец периода")
            self.splitter_5.setMaximumWidth(9999)
            self.splitter_8.setMaximumWidth(9999)
            self.splitter_9.setMaximumWidth(9999)
            self.lineEdit_2.setPlaceholderText("DD.MM.YYYY")
            self.lineEdit_3.setPlaceholderText("DD.MM.YYYY")
        else:
            self.label_5.setText("Заказчик")
            self.splitter_5.setMaximumWidth(9999)
            self.splitter_8.setMaximumWidth(0)
            self.splitter_9.setMaximumWidth(0)

    def text_changed(self):
        checked_but = [i for i in range(5) if self.radio_butts[i].isChecked()][0]
        text_1 = self.lineEdit_1.text()
        text_2 = self.lineEdit_2.text()
        text_3 = self.lineEdit_3.text()
        post = ""
        if checked_but == 3:
            text_2 = self.transform_data(text_2)
            text_3 = self.transform_data(text_3)
            if text_1:
                post += "WHERE ЗАКАЗЫ.заказчик='{}'".format(text_1)
                if text_2 or text_3:
                    if not text_2:
                        text_2 = "1970-01-01 "
                    elif not text_3:
                        text_3 = "2030-01-01 "
                    post += " AND (ЗАКАЗЫ.дата BETWEEN #{} 00:00:00# AND #{} 00:00:00#)".format(text_2, text_3)
            elif text_2 or text_3:
                if not text_2:
                    text_2 = "1970-01-01 "
                elif not text_3:
                    text_3 = "2030-01-01 "
                post = "WHERE ЗАКАЗЫ.дата BETWEEN #{} 00:00:00# AND #{} 00:00:00#".format(text_2, text_3)
            self.textEdit.setText(self.queries[checked_but].format(post))
        elif checked_but == 4:
            if text_1:
                post += "WHERE(((ГОТОВАЯ_ПРОДУКЦИЯ.заказчик)='{}'))".format(text_1)
        self.textEdit.setText(self.queries[checked_but].format(post))

    def connect_db(self):
        file_name = QFileDialog.getOpenFileName(self, "Выберите файл БД", '', "DB(*.mdb *.accdb)")[0]
        if file_name:
            self.db_manager.connect(file_name)
            self.fields = self.db_manager.get_db_fields()
            self.tables = self.db_manager.get_db_tables()
            self.field_types = self.db_manager.get_db_field_types()
            self.type_act.setCurrentIndex(0)
            self.table_act.addItems(self.tables)
            self.table_act.setCurrentIndex(0)
            self.table_act.setEnabled(True)
            self.type_act.setEnabled(True)
            self.create_but.setEnabled(True)
            self.first_rec.setEnabled(True)
            self.last_rec.setEnabled(True)
            self.tabWidget.setEnabled(True)
            self.resize_to_type(90, 43, 0, 0, 43)
            self.query_type_pos_changed()
            model = self.db_manager.model
            rowTotalWidth = 0
            for i in range(model.columnCount()):
                if self.field_types[0][i] == "CHAR":
                    self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
                    rowTotalWidth += self.table.horizontalHeader().sectionSize(i)
                    self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
                rowTotalWidth += self.table.horizontalHeader().sectionSize(i)
            rowTotalHeight = 0
            for i in range(10 if model.rowCount() > 10 else model.rowCount()):
                if not self.table.verticalHeader().isSectionHidden(i):
                    rowTotalHeight += self.table.verticalHeader().sectionSize(i)
            self.table.setMinimumSize(rowTotalWidth + 50,
                                      self.table.horizontalHeader().height() + rowTotalHeight + 40)
            self.import_action.setEnabled(False)
            self.setWindowTitle("Работа с БД - " + file_name)
            self.statusBar().showMessage("БД загружена! Списки таблиц и полей импортированы из БД!")

    def set_data_table(self, table, query, labels, temp_list, stretch=None):
        if not temp_list and query:
            not_err = query.first()
            while not_err and query.isValid():
                temp_list.append(tuple(query.value(i) for i in range(len(labels))))
                query.next()
        if table:
            new_model = QStandardItemModel(self)
            len_labels = len(labels)
            len_temp_list = len(temp_list)
            range_temp_list = range(len_temp_list)
            new_model.setRowCount(len_temp_list)
            new_model.setColumnCount(len_labels)
            for i in range_temp_list:
                for j in range(len_labels):
                    if isinstance(temp_list[i][j], QDateTime):
                        new_model.setItem(i, j, QStandardItem(temp_list[i][j].toString("dd.MM.yyyy")))
                    else:
                        new_model.setItem(i, j, QStandardItem(str(temp_list[i][j])))
            new_model.setHorizontalHeaderLabels(labels)
            table.setModel(new_model)
            for i in (range(len_labels) if not stretch else stretch):
                if not temp_list or isinstance(temp_list[0][i], str):
                    table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            if stretch:
                for i in range_temp_list:
                    table.verticalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
        return tuple(temp_list)

    def transform_data(self, data):
        try:
            data = data.split(".")
            data[0], data[1] = data[1], data[0]
            for i in range(2):
                if data[i][0] == "0":
                    data[i] = data[i][1]
            return "/".join(data)
        except Exception:
            return ""

    def exec_query_button(self):
        try:
            self.statusBar().showMessage("")
            if self.create_but.text() == "Сбросить результаты":
                self.create_but.setText("Выполнить запрос")
                index = self.type_act.currentIndex()
                self.table_pos_changed()
                self.type_act.setCurrentIndex(index)
                return
            cur_tab = self.table_act.currentText()
            cur_tab_index = self.table_act.currentIndex()
            tab_fields = self.fields[self.table_act.currentText()]
            count_fields = len(tab_fields)
            cur_values = [self.lines_cur[i].text() for i in range(count_fields)]
            cur_type = self.type_act.currentIndex()
            if cur_type == INSERT_TYPE:
                if not self.db_manager.query_insert(cur_tab_index, cur_tab, tab_fields,
                                                    cur_values):
                    self.statusBar().showMessage("Пустой запрос на добавление записи проигнорирован!")
                    return
            elif cur_type == DELETE_TYPE:
                if not self.db_manager.query_delete(cur_tab_index, cur_tab, tab_fields,
                                                    cur_values, self.lines_cond):
                    self.statusBar().showMessage("Пустой запрос на удаление проигнорирован!")
                    return
            elif cur_type == SELECT_TYPE:
                result = self.db_manager.query_select(cur_tab_index, cur_tab, tab_fields,
                                                      cur_values, self.lines_cond)
                if result:
                    self.set_data_table(self.table, result, tab_fields, [])
                    self.create_but.setText("Сбросить результаты")
                    self.statusBar().showMessage("Запрос выполнен!")
                else:
                    self.statusBar().showMessage("Пустой запрос на выборку проигнорирован!")
                return
            elif cur_type == UPDATE_TYPE:
                if not self.db_manager.query_update(cur_tab_index, cur_tab, tab_fields,
                                                    cur_values, self.lines_cond, self.lines_mod):
                    self.statusBar().showMessage("Пустой запрос на модификацию данных проигнорирован!")
                    return
            self.db_manager.query.finish()
            err = self.db_manager.last_err
            if not err:
                self.table_pos_changed()
                self.statusBar().showMessage("Запрос выполнен!")
            else:
                self.statusBar().showMessage(self.db_manager.last_err)
        except Exception as e:
            print(e)

    def exec_spec_query_button(self):
        try:
            self.create_report_but.setEnabled(True)
            self.table_query_res.clearSpans()
            self.table_query_res.setModel(QStandardItemModel(self))
            checked_but = [i for i in range(5) if self.radio_butts[i].isChecked()][0]
            result, indexes = [], []
            tab_fields = self.fields_query[checked_but]
            result = self.set_data_table(
                self.table_query_res, self.db_manager.exec(self.textEdit.toPlainText()), tab_fields, [])
            for field in range(len(tab_fields)):
                for record in range(1, len(result)):
                    if indexes:
                        if result[record][field] == result[record - 1][field]:
                            indexes[-1][2] += 1
                        else:
                            indexes.append([record, field, 1, 1])
                    else:
                        indexes.append([record - 1, field,
                                        2 if result[record][field] == result[record - 1][field] else 1,
                                        1])
            span_indexes = indexes.copy()
            for index in indexes:
                if index[2] != 1:
                    self.table_query_res.setSpan(*index)
                else:
                    span_indexes.remove(index)
            model = self.table_query_res.model()
            model.insertRow(model.rowCount(), QModelIndex())
            end_i = model.index(model.rowCount() - 1, 0)
            model.setData(end_i, self.postfixes[checked_but] + str(len(result)), QtCore.Qt.DisplayRole)
            self.table_query_res.setSpan(model.rowCount() - 1, 0, 1, 8)
            model.setData(end_i, QtCore.Qt.AlignVCenter | QtCore.Qt.AlignRight, QtCore.Qt.TextAlignmentRole)
            self.statusBar().showMessage("Запрос выполнен!")
            return checked_but, result, tab_fields, sorted(span_indexes)
        except Exception as e:
            print(e)
            self.statusBar().showMessage("Ошибка! Запрос не выполнен!")
            self.table_query_res.clearSpans()
            self.table_query_res.reset()
            self.open_report_but.setEnabled(False)
            self.create_report_but.setEnabled(False)

    def create_report(self):
        self.open_report_but.setEnabled(True)
        try:
            index, data, fields, spans = self.exec_spec_query_button()
            file_name, file_type = QFileDialog.getSaveFileName(
                self, "Сохранить отчет как", "report_{}".format(datetime.now().strftime("%H%M%S_%d%m%y")),
                "pdf (*.pdf);; MSWord (*.doc);; MSExcel (*.xlsx)")
            if file_name:
                report = Report(index, data, fields, self.postfixes)
                report.make_html(spans)
                report.create_report_file(file_type, file_name)
                self.statusBar().showMessage("Отчет создан и помещен в файл {}!".format(file_name))
                self.last_report_path = file_name
        except Exception as e:
            print(e)

    def init_controls(self):
        for i in range(1, 9):
            eval("self.lines_cur.append(self.field_" + str(i) + ")")
            eval("self.lines_cond.append(self.cond_" + str(i) + ")")
            eval("self.lines_mod.append(self.mod_" + str(i) + ")")
        for i in range(1, 6):
            eval("self.radio_butts.append(self.radioButton_" + str(i) + ")")
        self.import_action = QAction("Подключить базу данных")
        self.guide_action = QAction("Справка")
        self.about_action = QAction("О программе")
        self.exit_action = QAction("Выход")
        self.menubar.addAction(self.import_action)
        self.menubar.addAction(self.guide_action)
        self.menubar.addAction(self.about_action)
        self.menubar.addAction(self.exit_action)
        self.import_action.triggered.connect(self.connect_db)
        self.guide_action.triggered.connect(self.open_help)
        self.about_action.triggered.connect(lambda: self.statusBar().showMessage(
            "Приложение создано в учебных целях при выполнении курсовой работы по БД. Выполнил: ст. гр. ИПЗ-18-2 "
            "Комаров С.И."))
        self.exit_action.triggered.connect(self.close)
        self.table_act.currentTextChanged.connect(self.table_pos_changed)
        self.type_act.currentTextChanged.connect(self.query_type_pos_changed)
        self.create_but.clicked.connect(self.exec_query_button)
        self.first_rec.clicked.connect(self.table.clearSelection)
        self.first_rec.clicked.connect(lambda: self.table.selectRow(0))
        self.last_rec.clicked.connect(lambda: self.table.clearSelection())
        self.last_rec.clicked.connect(lambda: self.table.selectRow(self.db_manager.get_model().rowCount() - 1))
        self.radioButton_1.clicked.connect(lambda: self.radio_pos_changed(0))
        self.radioButton_2.clicked.connect(lambda: self.radio_pos_changed(1))
        self.radioButton_3.clicked.connect(lambda: self.radio_pos_changed(2))
        self.radioButton_4.clicked.connect(lambda: self.radio_pos_changed(3))
        self.radioButton_5.clicked.connect(lambda: self.radio_pos_changed(4))
        self.lineEdit_1.textChanged.connect(lambda: self.text_changed())
        self.lineEdit_2.textChanged.connect(lambda: self.text_changed())
        self.lineEdit_3.textChanged.connect(lambda: self.text_changed())
        self.tabWidget.currentChanged.connect(self.tab_window_changed)
        self.exec_spec_query_but.clicked.connect(self.exec_spec_query_button)
        self.create_report_but.clicked.connect(self.create_report)
        self.open_report_but.clicked.connect(lambda: system("start " + self.last_report_path.replace('\\', '/')))

    def open_help(self):
        path_help_fl = abspath(curdir).replace("\\", "\\\\") + "\\\\user_guide.doc"
        system("start {}".format(path_help_fl))
        self.statusBar().showMessage("Справка открыта!")

    def closeEvent(self, event):
        if self.db_manager.isOpen():
            self.db_manager.close_db()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
