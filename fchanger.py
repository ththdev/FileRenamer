# -*- coding: utf-8 -*- 
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from os import *
import pandas as pd
import numpy as np
import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


ui = resource_path('fchanger.ui')
ui_class = uic.loadUiType(ui)[0]
print(ui_class)

class WindowClass(QMainWindow, ui_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.loadButton.clicked.connect(self.load)
        self.saveButton.clicked.connect(self.write)
        self.loadExcelButton.clicked.connect(self.load_excel)
        self.applyButton.clicked.connect(self.apply)
        self.recoveryButton.clicked.connect(self.recovery)

    def load(self):
        fname = QFileDialog.getExistingDirectory(self)

        if fname:
            self.tableWidget.setRowCount(0);
            
            for (path, dir, files) in walk(fname):
                if len(files) > 0:
                    for i, file in enumerate(files):
                        ext = path.split(file)[-1]
                        self.tableWidget.insertRow(i)
                        self.tableWidget.setItem(i, 0, QTableWidgetItem(str(path)))
                        self.tableWidget.setItem(i, 1, QTableWidgetItem(str(file)))
    
    def write(self):
        path = QFileDialog.getSaveFileName(self, '엑셀로 저장하기', 'scan', 'Excel 통합 문서(*.xlsx)')
        
        if path[0]:
            rowCount = self.tableWidget.rowCount()
            count = 0
            col1 = []
            col2 = []

            while count < rowCount:
                col1.append(self.tableWidget.item(count, 0).text())
                col2.append(self.tableWidget.item(count, 1).text())
                count = count + 1
            
            temp_df = pd.DataFrame(columns=["경로", "변경전 파일이름", '변경할 파일이름', '변경완료'])
            temp_df['경로'] = col1
            temp_df['변경전 파일이름'] = col2
            try:
                temp_df.to_excel(path[0], index=False)
            except:
                msg = QMessageBox()
                msg.setWindowTitle("경고")
                msg.setText(f"{path[0]} 피일을 닫고 다시 저장해주세요")
                msg.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)
                result = msg.exec_()
    
    def load_excel(self):
        path = QFileDialog.getOpenFileName(self, '엑셀파일 불러오기', 'scan.xlsx', 'Excel 통합 문서(*.xlsx)')
        
        if path[0]:
            data = pd.read_excel(path[0]).fillna('')
            col1 = np.array(data['경로'])
            col2 = np.array(data['변경전 파일이름'])
            col3 = np.array(data['변경할 파일이름'])
            
            count = 0
            while count < len(data):
                self.tableWidget.insertRow(count)
                self.tableWidget.setItem(count, 0, QTableWidgetItem(str(col1[count])))
                self.tableWidget.setItem(count, 1, QTableWidgetItem(str(col2[count])))
                self.tableWidget.setItem(count, 2, QTableWidgetItem(str(col3[count])))
                count = count + 1

    def apply(self):
        rowCount = self.tableWidget.rowCount()
        i = 0
        while i < rowCount - 1:
            root = self.tableWidget.item(i, 0).text()
            old_name = self.tableWidget.item(i, 1).text()
            new_name = self.tableWidget.item(i, 2)
            if new_name:
                new_name = new_name.text()
                old = path.join(root, old_name)
                new = path.join(root, new_name)
                if new_name != "":
                    if new_name != old_name:
                        rename(old, new)
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("변경완료"))
            i = i + 1

    def recovery(self):
        rowCount = self.tableWidget.rowCount()
        i = 0
        while i < rowCount - 1:
            root = self.tableWidget.item(i, 0).text()
            old_name = self.tableWidget.item(i, 1).text()
            new_name = self.tableWidget.item(i, 2)
            if new_name:
                new_name = new_name.text()
                old = path.join(root, old_name)
                new = path.join(root, new_name)
                if new_name != "":
                    if new_name != old_name:
                        rename(new, old)
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("복구완료"))
            i = i + 1
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    myWindow = WindowClass()

    myWindow.show()

    app.exec_()