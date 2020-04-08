import os
import sys
import re
import xlrd
import win32com
from win32com.client import Dispatch
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QWidget, QApplication, QGroupBox, QPushButton, QLabel, QHBoxLayout, QVBoxLayout, \
    QGridLayout, QLineEdit, QTextEdit, QFileDialog, QMessageBox, QMainWindow


def get_tags(doc):
    text = ''
    for para in doc.Paragraphs:
        text += para.Range.text
    for table in doc.Tables:
        for row in table.Rows:
            for cell in row.Cells:
                text += cell.Range.text
    return re.findall('S[0-9]+[A-Z][0-9]+', text)


def split_tag(tag):
    match = re.match(r'S(\d+)([A-Z])(\d+)', tag)
    return int(match.group(1)) - 1, ord(match.group(2)) - ord('A'), int(match.group(3)) - 1 if match else None


def get_tag_values1(tags):
    tag_value_dict = {}
    for tag in tags:
        s, c, r = split_tag(tag)
        tag_value_dict[tag] = "S{}{}{}".format(s + 2, c, r)
    return tag_value_dict


def get_tag_values(book, tags):
    tag_value_dict = {}
    for tag in tags:
        s, c, r = split_tag(tag)
        sheet = book.sheets()[s]
        nrows = sheet.nrows
        ncols = sheet.ncols
        if r < nrows and c < ncols:
            tag_value_dict[tag] = sheet.cell(r, c).value
        else:
            tag_value_dict[tag] = None
    return tag_value_dict


def replace_doc(word_app, tag_value_dict):
    # xlApp.Selection.Find.ClearFormatting()
    # xlApp.Selection.Find.Replacement.ClearFormatting()
    # (string--搜索文本,
    # True--区分大小写,
    # True--完全匹配的单词，并非单词中的部分（全字匹配）,
    # True--使用通配符,
    # True--同音,
    # True--查找单词的各种形式,
    # True--向文档尾部搜索,
    # 1,
    # True--带格式的文本,
    # new_string--替换文本,
    # 2--替换个数（全部替换）
    result = ""
    for tag, value in tag_value_dict.items():
        if value:
            print(tag, ":", value)
            word_app.Selection.Find.Execute(tag, False, False, False, False, False, True, 1, True, value, 2)
        else:
            result += "从数据文件中找不到:{}\n".format(tag)
    return result


def show_msg(title, content, icon=3):
    box = QMessageBox(QMessageBox.Question, title, content)
    box.addButton('确定', QMessageBox.YesRole)
    box.setIcon(icon)
    box.exec()


class ReportGenerator(QWidget):

    def __init__(self):
        super(ReportGenerator, self).__init__()

        self.gridGroupBox = QGroupBox("基本参数")
        self.formGroupBox = QGroupBox("报告结果")

        self.template_label = QLabel('报告模板：')
        self.template_text = QLineEdit()
        self.template_text.setDisabled(True)
        self.template_btn = QPushButton('选择报告模板…')

        self.data_label = QLabel('数据文件：')
        self.data_text = QLineEdit()
        self.data_text.setDisabled(True)
        self.data_btn = QPushButton('选择数据文件…')

        self.submit_btn = QPushButton('生成报告')
        self.submit_btn.setStyleSheet("QPushButton{padding:20px 4px}")
        self.res_teatarea = QTextEdit()
        self.init_ui()

    def init_ui(self):
        self.create_grid_group_box()
        self.create_form_group_box()
        main_layout = QVBoxLayout()
        hbox_layout = QHBoxLayout()
        hbox_layout.addWidget(self.gridGroupBox)
        main_layout.addLayout(hbox_layout)
        main_layout.addWidget(self.formGroupBox)
        self.setLayout(main_layout)

    def create_grid_group_box(self):
        layout = QGridLayout()
        self.template_btn.clicked.connect(self.select_template_file)
        self.data_btn.clicked.connect(self.select_data_file)
        self.submit_btn.clicked.connect(self.process_win32)
        layout.setSpacing(10)
        layout.addWidget(self.template_label, 1, 0)
        layout.addWidget(self.template_text, 1, 1)
        layout.addWidget(self.template_btn, 1, 2)
        layout.addWidget(self.data_label, 2, 0)
        layout.addWidget(self.data_text, 2, 1)
        layout.addWidget(self.data_btn, 2, 2)
        layout.addWidget(self.submit_btn, 3, 0, 1, 3)
        layout.setColumnStretch(1, 10)
        self.gridGroupBox.setLayout(layout)
        self.setWindowTitle('报告生成器')

    def create_form_group_box(self):
        layout = QGridLayout()
        layout.addWidget(self.res_teatarea, 1, 0)
        self.formGroupBox.setLayout(layout)

    def select_template_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "请选择您要打开的文件", filter="*.docx")
        if len(file_name) > 0:
            if os.path.exists(file_name):
                self.template_text.setText(file_name)
            else:
                show_msg('错误', '您选择的文件不存在，请重新选择！')

    def select_data_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "请选择您要打开的文件", filter="*.xls")
        if len(file_name) > 0:
            if os.path.exists(file_name):
                self.data_text.setText(file_name)
            else:
                show_msg('错误', '您选择的文件不存在，请重新选择！')

    def process_win32(self):
        self.res_teatarea.moveCursor(QTextCursor.End)
        template_file = self.template_text.text()
        if not template_file:
            show_msg('错误', '您还没有选择报告模板')
            return False
        else:
            self.res_teatarea.insertPlainText("报告模板:{}\n".format(template_file))

        data_file = self.data_text.text()
        if not data_file:
            show_msg('错误', '您还没有选择数据文件')
            return False
        else:
            self.res_teatarea.insertPlainText("数据文件:{}\n".format(data_file))

        new_file = template_file[:-5]
        new_file += "_new.docx"
        self.res_teatarea.insertPlainText("结果文件:{}\n".format(new_file))

        word_app = win32com.client.Dispatch('Word.Application')
        word_app.Visible = 0
        word_app.DisplayAlerts = 0
        doc = word_app.Documents.Open(template_file)
        tags = get_tags(doc)

        book = xlrd.open_workbook(data_file)
        tag_value_dict = get_tag_values(book, tags)
        result = replace_doc(word_app, tag_value_dict)
        doc.SaveAs(new_file)
        word_app.Documents.Close()
        word_app.Quit()

        self.res_teatarea.insertPlainText("报告完成完成\n")
        self.res_teatarea.insertPlainText(result)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = ReportGenerator()
    ui.show()
    sys.exit(app.exec_())
