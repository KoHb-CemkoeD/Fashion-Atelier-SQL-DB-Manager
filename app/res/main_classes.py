from datetime import datetime
from os import path, curdir, remove, rmdir, walk
from threading import Thread
from pdfkit import from_string

from PyQt5.QtCore import QDateTime
from win32com.client.dynamic import Dispatch
from PyQt5.QtSql import QSqlDatabase, QSqlRelationalTableModel, QSqlQuery


class DBManager:
    def __init__(self):
        self._acces_field_types = {10: "CHAR", 16: "DATE", 2: "LONG INT", 6: "DOUBLE"}
        self.db_path = self.query = self.transform_data = self.model = None
        self.fields = {}
        self.tables, self.field_types, self.statistics = [], [], []
        self.last_err = ""
        self.data_base = QSqlDatabase()

    def connect(self, db_path):
        self.db_path = db_path
        self.data_base = QSqlDatabase.addDatabase("QODBC")
        self.data_base.setDatabaseName(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DNS='';DBQ=" + self.db_path)
        self.data_base.open()
        self.fields = {
            tab: [self.data_base.record(tab).fieldName(field)
                  for field in range(self.data_base.record(tab).count())]
            for tab in self.data_base.tables() if not tab.startswith("~")}
        self.tables = list(self.fields.keys())
        self.field_types = [
            [self._acces_field_types[self.data_base.record(tab).field(field).type()] for field in self.fields[tab]]
            for tab in self.tables]
        self.model = QSqlRelationalTableModel()
        self.query = QSqlQuery()

    def get_db_fields(self):
        return self.fields

    def get_db_tables(self):
        return self.tables

    def get_db_field_types(self):
        return self.field_types

    def isOpen(self):
        return self.data_base.isOpen()

    def close_db(self):
        self.data_base.close()

    def get_model(self):
        return self.model

    def get_model_by_index(self, index):
        self.model.setTable(self.tables[index])
        self.model.select()
        return self.model

    def exec(self, query_text, to_statistic=True):
        self.query.exec(query_text)
        err = self.query.lastError().text()
        status = ""
        if not err:
            status = ""
        elif "duplicate values" in err:
            status += "Ошибка! Обнаружены совпадения данных в ключевых полях!"
        elif "validation rule" in err:
            status += "Ошибка! Неверный ввод значения для поля " + \
                      err[
                      err.find("set for") + 7:err.rfind("'") + 1] + " ! Необходимые условия: " + \
                      + err[err.find("rule") + 7:err.rfind("' set") + 1]
        elif "Data type mismatch" in err or "Unable to execute" in err:
            status += "Ошибка! Обнаружено несоответстивие типов входных данных!"
        elif "Syntax error" in err:
            status += "Ошибка! Неверный ввод условия!"
        else:
            status += "Ошибка! Запрос не выполнен!"
        if to_statistic:
            self.statistics.append(
                [query_text, datetime.now().strftime("%H:%M:%S %d.%m.%Y"),
                 (status + "\n\tПодробнее: " + err) if status else "Успех!"])
        self.last_err = status
        return self.query

    def query_select(self, cur_tab_index, cur_tab, tab_fields, cur_values, lines_cond):
        cond_val, pref_f = "", ""
        for ind_field, field in enumerate(tab_fields):
            if pref_f:
                pref_f += ", "
            pref_f += " {}.[{}]".format(cur_tab, field)
            if lines_cond[ind_field].text():
                if cur_values[ind_field]:
                    type_field = self.field_types[cur_tab_index][ind_field]
                    val = ("'{}'".format(cur_values[ind_field]) if type_field == "CHAR" else
                           "#{}#".format(self.transform_data(cur_values[ind_field])) if type_field == "DATE" else
                           cur_values[ind_field])
                    if cond_val:
                        cond_val += " and "
                    cond_val += ("((" + cur_tab + ".[" + field + "])" + lines_cond[
                        ind_field].text() + val + ")")
        if cond_val:
            sql_text = "SELECT {} FROM {} WHERE ({})".format(pref_f, cur_tab, cond_val)
            return self.exec(sql_text)

    def query_insert(self, cur_tab_index, cur_tab, tab_fields, cur_values):
        corrected_field_values, empty = [], True
        for ind_field, field in enumerate(tab_fields):
            type_field = self.field_types[cur_tab_index][ind_field]
            value = cur_values[ind_field]
            if value:
                empty = False
                corrected_field_values.append(
                    "'{}'".format(value) if type_field == "CHAR" or type_field == "DATE" else cur_values[
                        ind_field])
            else:
                corrected_field_values.append(
                    "''" if type_field == "CHAR" else '#00:00:00#' if type_field == "DATE" else "0.0")
        if not empty:
            sql_text = "INSERT INTO {}({}) VALUES({})".format(
                cur_tab, ", ".join(tab_fields), ",".join(corrected_field_values))
            self.exec(sql_text)
            return True

    def query_delete(self, cur_tab_index, cur_tab, tab_fields, cur_values, lines_cond):
        cond_val, pref_f = "", ""
        for ind_field, field in enumerate(tab_fields):
            cond = lines_cond[ind_field].text()
            if cur_values[ind_field] and cond:
                type_field = self.field_types[cur_tab_index][ind_field]
                field_cond = (field if type_field == "CHAR" or type_field == "DATE" else "[{}]".format(field))
                val = ("'{}'".format(cur_values[ind_field]) if type_field == "CHAR" else
                       "#{}#".format(
                           self.transform_data(cur_values[ind_field])) if type_field == "DATE" else
                       cur_values[ind_field])
                if cond_val:
                    cond_val += " AND "
                cond_val += ("(({}.{}) {} {})".format(cur_tab, field_cond, cond, val))
                if pref_f:
                    pref_f += ", "
                pref_f += cur_tab + "." + field
        if cond_val:
            sql_text = "DELETE {} AS Выражение1, {} FROM {} WHERE (({}))".format(
                cur_tab, pref_f, cur_tab, cond_val)
            self.exec(sql_text)
            return True

    def query_update(self, cur_tab_index, cur_tab, tab_fields, cur_values, lines_cond, lines_mod):
        cond_val, pref_f, new_val, type_field = "", "", "", ""
        for ind_field, field in enumerate(tab_fields):
            cond = lines_cond[ind_field].text()
            new = lines_mod[ind_field].text()
            if cur_values[ind_field] and cond:
                type_field = self.field_types[cur_tab_index][ind_field]
                field_cond = (field if type_field == "CHAR" or type_field == "DATE" else "[{}]".format(field))
                val = ("'{}'".format(cur_values[ind_field]) if type_field == "CHAR" else
                       "#{}#".format(self.transform_data(cur_values[ind_field])) if type_field == "DATE"
                       else cur_values[ind_field])
                if cond_val:
                    cond_val += " and "
                cond_val += ("(({}.{}) {} {})".format(cur_tab, field_cond, cond, val))
            if new:
                new_val = ("'{}'".format(new) if type_field == "CHAR" else
                           "#{}#".format(new.replace(".", "/")) if type_field == "DATE" else
                           new)
                if pref_f:
                    pref_f += ", "
                pref_f += " {}.{} = {}".format(cur_tab, field, new_val)
        if cond_val:
            sql_text = \
                "UPDATE {} SET {} WHERE (({}))".format(cur_tab, pref_f, cond_val)
            self.exec(sql_text)
            return True
        else:
            return False

    def exec_and_get_data(self, sql_text):
        temp_list = []
        self.query.exec(sql_text)
        if not temp_list:
            not_err = self.query.first()
            while not_err and self.query.isValid():
                temp_list.append(self.query.value(0))
                self.query.next()
        self.query.finish()
        return tuple(temp_list)


class Report:
    def __init__(self, index, report_data, fields, postfixes):
        self.index, self.report_data, self.fields, self.postfixes = index, report_data, fields, postfixes
        self.report_doc = self.report_pdf = self.file_name = ""
        self.names = ["Отчет по моделям", "Отчет по заказчикам", "Отчет по закройщикам",
                      "Отчет по заказам за период и по заказчикам", "Отчет готовой продукции по заказчикам"]
        with open("res/report_pdf.html") as report_pdf_header:
            self.report_pdf_header = report_pdf_header.read()
        with open("res/report_doc.html") as report_doc_header:
            self.report_doc_header = report_doc_header.read()

    def _report_doc(self):
        with open("res/temp.html", "w") as temp_report:
            temp_report.write(self.report_doc)
        wordApp = Dispatch("Word.Application")
        wordDoc = wordApp.Documents.Open(
            path.abspath(curdir).replace("\\", "\\\\") + "\\\\res\\\\temp.html")
        wordDoc.SaveAs(self.file_name.replace(r":/", r":\\").replace(r"/", r"\\"))
        wordDoc.Close()
        wordApp.Quit()
        temp_dir = self.file_name[:self.file_name.rindex(".")] + ".files"
        for root, dirs, files in walk(temp_dir, topdown=False):
            for name in files:
                remove(path.join(root, name))
            for name in dirs:
                rmdir(path.join(root, name))
        rmdir(temp_dir)

    def _report_excel(self):
        excel_app = Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.ScreenUpdating = False
        work_book = excel_app.Workbooks.Add()
        work_sheet = work_book.ActiveSheet
        for i, field in enumerate(self.fields):
            work_sheet.Cells(1, i + 1).value = field
        for i, row in enumerate(self.report_data):
            for j, elem in enumerate(row):
                work_sheet.Cells(i + 2, j + 1).value = \
                    elem.toString("dd.MM.yyyy") if isinstance(elem, QDateTime) else elem
        work_sheet.Columns.AutoFit()
        work_sheet.ListObjects.Add().TableStyle = "TableStyleLight8"
        work_sheet.Name = self.names[self.index][:31]
        work_book.SaveAs(self.file_name.replace("/", "\\"))
        work_book.Close()
        excel_app.Quit()

    def make_html(self, spans):
        len_head = str(len(self.fields))
        text_report = ""
        if not spans:
            text_report += "<tr>" + "</tr>\n<tr>".join(
                ["<td rowspan='1'>" + "</td><td rowspan='1'>".join(
                    [j.toString("dd.MM.yyyy") if isinstance(j, QDateTime) else str(j)
                     for j in i]) + "</td>" for i in self.report_data]) + "</tr>"
        else:
            prev_data = ["" for i in range(len(self.fields))]
            for ind_rec, record in enumerate(self.report_data):
                text_report += "<tr>\n"
                for ind_data, data in enumerate(record):
                    data = data.toString("dd.MM.yyyy") if isinstance(data, QDateTime) else str(data)
                    if spans and spans[0][0] == ind_rec and spans[0][1] == ind_data:
                        prev_data[ind_data] = data
                        text_report += "\t<td rowspan='" + str(spans[0][2]) + "'>" + data + "</td>\n"
                        spans.pop(0)
                    else:
                        if data != prev_data[ind_data]:
                            text_report += "\t<td rowspan='1'>" + data + "</td>\n"
                text_report += "</tr>\n"
        text_report += "\n\t\t<tr><td align='right' colspan='{}'>{} {}</td></tr>\n\t".format(
            len_head, self.postfixes[self.index], str(len(self.report_data)))
        self.report_doc = self.report_doc_header \
            .replace("len_head", len_head) \
            .replace("text_report", text_report) \
            .replace("report_name", self.names[self.index]) \
            .replace("report_header", "<th>" + "</th><th>".join(self.fields) + "</th>")
        self.report_pdf = self.report_pdf_header \
            .replace("len_head", len_head) \
            .replace("text_report", text_report) \
            .replace("report_name", self.names[self.index]) \
            .replace("report_header", "<th>" + "</th><th>".join(self.fields) + "</th>")

    def create_report_file(self, file_type, file_name):
        self.file_name = file_name
        if file_type == "pdf (*.pdf)":
            options = {'page-size': 'Letter', 'margin-top': '0.75in', 'margin-right': '0.75in',
                       'margin-bottom': '0.75in',
                       'margin-left': '0.75in', 'encoding': "UTF-8", 'footer-center': '[page]',
                       'cookie': [
                           ('cookie-name1', 'cookie-value1'),
                           ('cookie-name2', 'cookie-value2'),
                       ], 'no-outline': None
                       }
            create_report = lambda: from_string(self.report_pdf, file_name, options=options)
        elif file_type == "MSWord (*.doc)":
            create_report = self._report_doc
        elif file_type == "MSExcel (*.xlsx)":
            create_report = self._report_excel
        thread = Thread(target=create_report)
        thread.start()


class Person:
    editing_parameters = ["full_name", "last_name", "middle_name", "first_name",
                          "phone", "adress"]

    def __init__(self, full_name="- - -", phone="", adress=""):
        self._last_name, self._first_name, self._middle_name = full_name.split()
        self.full_name = full_name
        self.phone, self.adress = phone, adress

    def set(self, full_name, b_day, ident_code):
        self.__init__(full_name, b_day, ident_code)

    def edit(self, params={}):
        for i in params:
            if i in "last_name first_name middle_name":
                self.full_name = self.full_name.replace(eval('self._' + i), params[i])
                exec('self._' + i + ' = "' + params[i] + '"')
            elif i in "b_day b_month b_year":
                self.full_b_day = self.full_b_day.replace(eval('self._' + i), params[i])
                exec('self._' + i + ' = "' + params[i] + '"')
            elif "full" in i:
                exec('self.' + i + ' = "' + params[i] + '"')
                temp_full_name = self.full_name.split()
                full_name = ["", "", ""]
                full_name[0:len(temp_full_name)] = temp_full_name
                self._last_name, self._first_name, self._middle_name = full_name
            else:
                self.ident_code = params[i]

    def __str__(self):
        return self.full_name + "; " + self.phone + "; " + self.adress


class Cutter(Person):
    operations = ["full_name", "last_name", "middle_name", "first_name",
                          "phone", "adress", "specialization", "work_expirience"]

    def __init__(self, full_name="- - -", phone="", adress="", specialization="", work_expirience=0):
        self.specialization, self.work_expirience = specialization, work_expirience
        super().__init__(full_name, phone, adress)

    def set(self, **kwargs):
        self.__init__(kwargs)

    def edit(self, params={}):
        for i in params:
            if "name" in i or "phone" in i or "adress " in i:
                Person.edit(self, params)
            elif i == "specialization":
                self.specialization = params[i]
            elif i == "work_expirience":
                self.work_expirience = params[i]

    def get(self, params):
        if "name" in params:
            return eval('self._' + params)
        elif "phone" in params:
            return self.phone
        elif "adress" in params:
            return self.adress

    def __str__(self):
        return self.full_name + "; " + self.phone + "; " + self.adress \
               + "; " + self.specialization + "; " + self.work_expirience


class Order:
    operations = ["number", "data", "name", "size",
                          "issued", "customer", "cutter"]

    def __init__(self, number=0, _data="", _customer="", _cutter=Cutter(), _name="", _size="", _issued=""):
        self._number, self._data, self._name, self._size, self._issued, = number, _data, _name, _size, _issued
        self._customer = _customer
        self._cutter = _cutter

    def set(self, **kwargs):
        self.__init__(kwargs)

    def edit(self, params={}):
        for i in params:
            eval('self._' + i + " = " + params[i])

    def get(self, params):
        return eval('self._' + params)

    def __str__(self):
        return "; ".join(str(eval('self._' + param) for param in self.editing_parameters))


class Model:
    operations = ["_model_number", "_purpose", "_specialization", "size"]

    def __init__(self, _model_number=0, _purpose="", _specialization="", _size=0.0):
        self._model_number, self._size = _model_number, _size
        self._purpose, self._specialization = _purpose, _specialization

    def set(self, about, specialization="", work_expirience=0):
        self.__init__(about, specialization, work_expirience)

    def edit(self, params={}):
        for i in params:
            eval('self._' + i + " = " + params[i])

    def get(self, params):
        return eval('self._' + params)

    def __str__(self):
        return "; ".join(str(eval('self._' + param) for param in self.operations))


class FinishedProduct:
    operations = ["code", "model", "size", "customer", "finished_data", "material"]

    def __init__(self, full_information):
        info = full_information.split(";")
        self.code, self._model, self._size, self._customer, self._finished_data, self._material = info

    def set(self, full_information):
        self.__init__(type, full_information)

    def get(self, params):
        return eval('self._' + params)

    def edit(self, params={}):
        for i in params:
            eval('self._' + i + " = " + params[i])

    def check(self, params={}):
        for i in params:
            if eval('self._' + i + ' != params[i]'):
                return False
        return True


class Purchase:
    operations = ["code", "data", "fabric_article", "length", "provider"]

    def __init__(self, full_information):
        info = full_information.split(";")
        self.code, self._data, self._fabric_article, self._length, self._provider = info

    def set(self, full_information):
        self.__init__(type, full_information)

    def get(self, params):
        return eval('self._' + params)

    def edit(self, params={}):
        for i in params:
            eval('self._' + i + " = " + params[i])

    def check(self, params={}):
        for i in params:
            if eval('self._' + i + ' != params[i]'):
                return False
        return True
