import os
import sys
import re
import xlrd
from hashlib import md5
from win32com.client import Dispatch
from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QGroupBox, QPushButton, QLabel, QHBoxLayout, QVBoxLayout, \
    QGridLayout, QLineEdit, QTextEdit, QFileDialog, QMessageBox, QMainWindow

key_word = r'XX审字[2020]'
tag_md5 = r'433dff137657d6f9971f75fef4877f08'
#key_word = r'国锐审字[2020]'
#tag_md5 = r'3953cf4ab9b5d6a25f715f62030b94ba'
percentage_tags = ['S4D27', 'S3H27', 'S3H28', 'S3H29', 'S3H30', 'S3H31']
form_tags = ['S2C04', 'S2C05', 'S2C06', 'S2C07', 'S2C08', 'S2C09', 'S2C10', 'S2C11', 'S2C12', 'S2C13', 'S2C14', 'S2C15',
             'S2C16', 'S2C17', 'S2C18', 'S2C19', 'S2C20', 'S2C21', 'S2C22', 'S2C23', 'S2C24', 'S2C25', 'S2C26', 'S2C27',
             'S2C28', 'S2C29', 'S2C30', 'S2C31', 'S2C32', 'S2C33', 'S2C34', 'S2C35', 'S2C36', 'S2C37', 'S2C38', 'S2C39',
             'S2C40']
company_name_tag = 'S1C14'


# c_type : 0 empty, 1 string, 2 number, 3 date, 4 boolean, 5 error, 6 blank
def format_value(tag, value, c_type):
    ret = None
    if 1 == c_type:
        ret = value.strip()
    elif 2 == c_type:
        if tag in percentage_tags:
            ret = '{:.2%}'.format(value)
        elif isinstance(value, float):
            ret = format(value, ',')
        else:
            ret = value
    return ret


def check_tag(tags):
    ret = False
    if tags:
        tag_str = ','.join(tags)
        # print(md5(tag_str.encode(encoding='UTF-8')).hexdigest())
        if tag_md5 == md5(tag_str.encode(encoding='UTF-8')).hexdigest():
            ret = True
    return ret


def get_tags(doc):
    text = ''
    for table in doc.Tables:
        for row in table.Rows:
            for cell in row.Cells:
                text += cell.Range()
    for para in doc.Paragraphs:
        text += para.Range()
    if key_word not in text:
        return None
    else:
        return re.findall('S[0-9]+[A-Z][0-9]+', text)


def split_tag(tag):
    match = re.match(r'S(\d+)([A-Z])(\d+)', tag)
    return int(match.group(1)) - 1, ord(match.group(2)) - ord('A'), int(match.group(3)) - 1 if match else None


def get_tag_values(book, tags):
    tag_value_dict = {}
    form_index = 1
    for tag in tags:
        s, c, r = split_tag(tag)
        sheet = book.sheets()[s]
        n_rows = sheet.nrows
        n_cols = sheet.ncols
        tag_value_dict[tag] = None
        if r < n_rows and c < n_cols:
            if tag in form_tags:
                if r'√' == sheet.cell(r, c).value:
                    tag_value_dict[tag] = r'{}.《{}》（{}）'.format(form_index, sheet.cell(r, 1).value.strip(),
                                                                sheet.cell(r, 0).value)
                    form_index += 1
            else:
                tag_value_dict[tag] = format_value(tag, sheet.cell(r, c).value, sheet.cell(r, c).ctype)
    return tag_value_dict


def delete_line(word_app, tag):
    # select a Text
    word_app.Selection.Find.Execute(tag)
    #  extend it to end
    word_app.Selection.EndKey(Unit=5, Extend=1)  # win32com.client.constants.wdLine, win32com.client.constants.wdExtend
    # check what has been selected
    word_app.Selection.Range()
    # and then delete it
    word_app.Selection.Delete()


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
        result += "{}:{}\n".format(tag, value)
        if value is None:
            if tag in form_tags:
                delete_line(word_app, tag)
                continue
            else:
                value = 0
        word_app.Selection.Find.Execute(tag, False, False, False, False, False, True, 1, True, value, 2)
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
        self.setWindowState(Qt.WindowMaximized)
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
        self.setWindowTitle('企业所得税审核报告生成器')

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

        word_app = Dispatch('Word.Application')
        word_app.Visible = 0
        word_app.DisplayAlerts = 0
        doc = word_app.Documents.Open(template_file)
        tags = get_tags(doc)

        if not check_tag(tags):
            show_msg('错误', '模版文件不符合规则！')
        else:
            book = xlrd.open_workbook(data_file, formatting_info=True)
            tag_value_dict = get_tag_values(book, tags)
            result = replace_doc(word_app, tag_value_dict)

            (file_path, _) = os.path.split(template_file)
            new_file = '{}\\2019年企业所得税审核报告及说明_{}.docx'.format(file_path, tag_value_dict[company_name_tag])
            doc.SaveAs(new_file)
            self.res_teatarea.insertPlainText("报告完成:{}\n".format(new_file))
            self.res_teatarea.insertPlainText(result)
        word_app.Documents.Close()
        word_app.Quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = ReportGenerator()
    ui.show()
    sys.exit(app.exec_())
