import pandas as pd
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import sys
import icon_rc


class Ui_choose_file(QtWidgets.QMainWindow):
    excel_dict = ''
    df_dict = pd.DataFrame()
    dict_num_dir = ''
    dict_sheet_name = ''
    get_path = ''
    excel_get = ''
    df_get = pd.DataFrame()
    get_num_dir = ''
    get_sheet_name = ''

    def setupUi(self, choose_file):
        choose_file.setObjectName("choose_file")
        choose_file.resize(750, 275)
        choose_file.setWindowIcon(QtGui.QIcon(':/ico/search.ico'))
        self.centralwidget = QtWidgets.QWidget(choose_file)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(50, 20, 640, 120))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_choose_dict = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_dict.setObjectName("label_choose_dict")
        self.horizontalLayout_4.addWidget(self.label_choose_dict)
        self.dict_dir = QtWidgets.QTextEdit(self.verticalLayoutWidget)
        self.dict_dir.setObjectName("dict_dir")
        self.horizontalLayout_4.addWidget(self.dict_dir)
        self.btn_choose_dict = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.btn_choose_dict.setObjectName("btn_choose_dict")
        self.btn_choose_dict.clicked.connect(self.choose_dict)
        self.horizontalLayout_4.addWidget(self.btn_choose_dict)
        self.verticalLayout_2.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_choose_dict_sheet_name = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_dict_sheet_name.setObjectName("label_choose_dict_sheet_name")
        self.horizontalLayout_6.addWidget(self.label_choose_dict_sheet_name)
        self.choose_dict_sheet_name = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.choose_dict_sheet_name.setObjectName("choose_dict_sheet_name")
        self.choose_dict_sheet_name.currentIndexChanged[str].connect(self.select_dict_sheet_name)
        self.horizontalLayout_6.addWidget(self.choose_dict_sheet_name)
        self.label_choose_dict_column = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_dict_column.setObjectName("label_choose_dict_column")
        self.horizontalLayout_6.addWidget(self.label_choose_dict_column)
        self.choose_dict_column = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.choose_dict_column.setObjectName("choose_dict_column")
        self.choose_dict_column.currentIndexChanged[str].connect(self.select_dict_column)
        self.horizontalLayout_6.addWidget(self.choose_dict_column)
        self.verticalLayout_2.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_choose_get = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_get.setObjectName("label_choose_get")
        self.horizontalLayout_5.addWidget(self.label_choose_get)
        self.get_dir = QtWidgets.QTextEdit(self.verticalLayoutWidget)
        self.get_dir.setObjectName("get_dir")
        self.horizontalLayout_5.addWidget(self.get_dir)
        self.btn_choose_get = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.btn_choose_get.setObjectName("btn_choose_get")
        self.btn_choose_get.clicked.connect(self.choose_get)
        self.horizontalLayout_5.addWidget(self.btn_choose_get)
        self.verticalLayout_2.addLayout(self.horizontalLayout_5)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_choose_get_sheet_name = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_get_sheet_name.setObjectName("label_choose_get_sheet_name")
        self.horizontalLayout_7.addWidget(self.label_choose_get_sheet_name)
        self.choose_get_sheet_name = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.choose_get_sheet_name.setObjectName("choose_get_sheet_name")
        self.choose_get_sheet_name.currentIndexChanged[str].connect(self.select_get_sheet_name)
        self.horizontalLayout_7.addWidget(self.choose_get_sheet_name)
        self.label_choose_get_column = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_choose_get_column.setObjectName("label_choose_get_column")
        self.horizontalLayout_7.addWidget(self.label_choose_get_column)
        self.choose_get_column = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.choose_get_column.setObjectName("choose_get_column")
        self.choose_get_column.currentIndexChanged[str].connect(self.select_get_column)
        self.horizontalLayout_7.addWidget(self.choose_get_column)
        self.verticalLayout_2.addLayout(self.horizontalLayout_7)
        self.btn_activate = QtWidgets.QPushButton(self.centralwidget)
        self.btn_activate.setGeometry(QtCore.QRect(275, 175, 200, 75))
        self.btn_activate.setObjectName("btn_activate")
        font = QtGui.QFont()
        font.setFamily('微软雅黑')
        font.setBold(True)
        font.setPointSize(24)
        self.btn_activate.setFont(font)
        self.btn_activate.clicked.connect(self.activate)
        choose_file.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(choose_file)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 750, 25))
        self.menubar.setObjectName("menubar")
        choose_file.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(choose_file)
        self.statusbar.setObjectName("statusbar")
        choose_file.setStatusBar(self.statusbar)

        self.retranslateUi(choose_file)
        QtCore.QMetaObject.connectSlotsByName(choose_file)

    def retranslateUi(self, choose_file):
        _translate = QtCore.QCoreApplication.translate
        choose_file.setWindowTitle(_translate("choose_file", "表格筛选小工具"))
        self.label_choose_dict.setText(_translate("choose_file", "请选择你的BOM表excel文件："))
        self.btn_choose_dict.setText(_translate("choose_file", "浏览"))
        self.label_choose_dict_sheet_name.setText(_translate("choose_file", "请选择工作表名："))
        self.label_choose_dict_column.setText(_translate("choose_file", "请选择物料号列名："))
        self.label_choose_get.setText(_translate("choose_file", "请选择待筛选的excel文件： "))
        self.btn_choose_get.setText(_translate("choose_file", "浏览"))
        self.label_choose_get_sheet_name.setText(_translate("choose_file", "请选择工作表名："))
        self.label_choose_get_column.setText(_translate("choose_file", "请选择物料号列名："))
        self.btn_activate.setText(_translate("choose_file", "运行"))

    def choose_dict(self):
        """选择BOM表"""
        self.dict_dir.clear()
        dict_path, dict_type = QFileDialog.getOpenFileName(self,
                                                           '选取文件',
                                                           '.\\',
                                                           'Excel Files (*.xlsx);;All Files (*)')
        if dict_path != '':
            print('已选择BOM表：' + dict_path)
            self.dict_dir.setText(dict_path)
            self.excel_dict = pd.ExcelFile(dict_path)
            # print(self.excel_dict.sheet_names)
            self.choose_dict_sheet_name.clear()
            self.choose_dict_sheet_name.addItems(self.excel_dict.sheet_names)

    def select_dict_sheet_name(self, value):
        """选择BOM表工作表"""
        if value != '':
            print('已选择BOM表工作表：' + value)
            self.dict_sheet_name = value
            self.df_dict = pd.read_excel(self.excel_dict, value)
            # print(self.df_dict.columns)
            self.choose_dict_column.clear()
            self.choose_dict_column.addItems(self.df_dict.columns)

    def select_dict_column(self, value):
        """选择BOM表中物料号所在列"""
        if value != '':
            print('已选择BOM表中物料号所在列：' + value)
            self.dict_num_dir = value

    def choose_get(self):
        """选择目标文件"""
        self.get_dir.clear()
        get_path, get_type = QFileDialog.getOpenFileName(self,
                                                         '选取文件',
                                                         '.\\',
                                                         'Excel Files (*.xlsx);;All Files (*)')
        if get_path != '':
            print('已选择目标文件：' + get_path)
            self.get_path = get_path
            self.get_dir.setText(get_path)
            self.excel_get = pd.ExcelFile(get_path)
            # print(self.excel_get.sheet_names)
            self.choose_get_sheet_name.addItems(self.excel_get.sheet_names)

    def select_get_sheet_name(self, value):
        """选择目标工作表"""
        if value != '':
            print('已选择目标工作表：' + value)
            self.get_sheet_name = value
            self.df_get = pd.read_excel(self.excel_get, value)
            # print(self.df_get.columns)
            self.choose_get_column.clear()
            # self.choose_get_column.addItem('请选择')
            self.choose_get_column.addItems(self.df_get.columns)

    def select_get_column(self, value):
        """选择目标表中物料号所在列"""
        if value != '':
            print('已选择目标表中物料号所在列:' + value)
            self.get_num_dir = value

    def activate(self):
        """运行核心功能"""
        if self.excel_dict == '' or self.excel_get == '':
            QMessageBox.warning(self, "提示", "请先完成文件选择！")
        else:
            df_dict = self.df_dict
            dict_num_dir = self.dict_num_dir
            df_get = self.df_get
            get_num_dir = self.get_num_dir
            index = 0
            df_res_get = pd.DataFrame(columns=df_get.columns.values)
            df_res_dict = pd.DataFrame(columns=df_dict.columns.values)
            # print(df_res)
            for num_get in df_get[get_num_dir]:
                for num_dict in df_dict[dict_num_dir]:
                    if num_dict == num_get:
                        index += 1
                        df_res_get.loc[index] = df_get[df_get[get_num_dir] == num_get].iloc[0]
                        df_res_dict.loc[index] = df_dict[df_dict[dict_num_dir] == num_dict].iloc[0]
                        break
            res_path = self.get_path[0:-5] + '_已筛选.xlsx'
            res_sheet_name_get = '目标表_' + self.get_sheet_name
            res_sheet_name_dict = 'BOM表_' + self.dict_sheet_name
            writer = pd.ExcelWriter(path=res_path)
            df_res_get.to_excel(excel_writer=writer, sheet_name=res_sheet_name_get, index=None)
            df_res_dict.to_excel(excel_writer=writer, sheet_name=res_sheet_name_dict, index=None)
            writer.save()
            print('筛选成功，已保存在目标表所在文件夹下！')
            QMessageBox.about(self, "提示", "筛选成功，已保存在目标表所在文件夹下！")


def show_MainWindow():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_choose_file()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    show_MainWindow()
