import sys

import sqlite3
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QTableWidgetItem


from mainwindow import Ui_MainWindow
from scripts.PortsExplorer.add_port import Ui_AddPort
from scripts.PortsExplorer.db import Database


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Main, self).__init__()
        self.window = Ui_MainWindow()
        self.window.setupUi(self)
        self.db = Database()
        self.entries = self.db.get_all()
        self.last_entries = self.entries
        dropdown1 = set()
        dropdown2 = set()
        for entry in self.entries:
            dropdown1.add(entry[1])
            dropdown2.add(entry[2])
            row_position = self.window.tableWidget.rowCount()
            self.window.tableWidget.insertRow(row_position)
            self.window.tableWidget.setItem(row_position, 0, QTableWidgetItem(entry[0]))
            self.window.tableWidget.setItem(row_position, 1, QTableWidgetItem(entry[1]))
            self.window.tableWidget.setItem(row_position, 2, QTableWidgetItem(entry[2]))
        for d in sorted(dropdown1):
            self.window.comboBox.addItem(d)
        for d in sorted(dropdown2):
            self.window.comboBox_2.addItem(d)
        self.window.lineEdit.textChanged.connect(self.get_text)
        self.window.tableWidget.horizontalHeader().setSectionResizeMode(1)
        self.window.tableWidget.sortItems(0, Qt.AscendingOrder)
        self.window.comboBox.currentTextChanged.connect(self.get_text)
        self.window.comboBox_2.currentTextChanged.connect(self.get_text)
        self.window.pushButton_2.clicked.connect(self.add_dialog)
        self.window.pushButton.clicked.connect(self.remove_entry)

    def remove_entry(self):
        rows = (set(index.row() for index in self.window.tableWidget.selectedIndexes()))
        for r in rows:
            item = self.window.tableWidget.item(r, 0).text()
            self.window.tableWidget.removeRow(r)
            self.db.remove(item)
        self.window.comboBox.currentTextChanged.disconnect()
        self.window.comboBox_2.currentTextChanged.disconnect()
        self.window.lineEdit.textChanged.disconnect()
        self.window.tableWidget.setRowCount(0)
        self.window.comboBox.clear()
        self.window.comboBox_2.clear()
        self.window.tableWidget.setSortingEnabled(False)
        self.window.comboBox.addItem("Select destination name")
        self.window.comboBox_2.addItem("Select destination code")
        self.window.lineEdit.setText('')
        self.entries = self.db.get_all()
        dropdown1 = set()
        dropdown2 = set()
        for entry in self.entries:
            dropdown1.add(entry[1])
            dropdown2.add(entry[2])
            row_position = self.window.tableWidget.rowCount()
            self.window.tableWidget.insertRow(row_position)
            self.window.tableWidget.setItem(row_position, 0, QTableWidgetItem(entry[0]))
            self.window.tableWidget.setItem(row_position, 1, QTableWidgetItem(entry[1]))
            self.window.tableWidget.setItem(row_position, 2, QTableWidgetItem(entry[2]))
        for d in sorted(dropdown1):
            self.window.comboBox.addItem(d)
        for d in sorted(dropdown2):
            self.window.comboBox_2.addItem(d)
        self.window.lineEdit.textChanged.connect(self.get_text)
        self.window.tableWidget.horizontalHeader().setSectionResizeMode(1)
        self.window.tableWidget.sortItems(0, Qt.AscendingOrder)
        self.window.comboBox.currentTextChanged.connect(self.get_text)
        self.window.comboBox_2.currentTextChanged.connect(self.get_text)
        self.window.pushButton_2.clicked.connect(self.add_dialog)
        self.window.tableWidget.setSortingEnabled(True)
        self.window.pushButton_3.clicked.connect(self.add_column)

    def add_column(self):
        pass

    def add_dialog(self):
        self.addport = QMainWindow()
        self.addport.ui = Ui_AddPort()
        self.addport.ui.setupUi(self.addport)
        self.addport.setWindowTitle('Add port')
        self.addport.ui.pushButton_2.clicked.connect(self.addport.close)
        self.addport.ui.pushButton.clicked.connect(self.add_this_port)
        self.addport.ui.pushButton.setDisabled(True)
        self.addport.ui.lineEdit.textChanged.connect(self.disable_if_empty)
        self.addport.ui.lineEdit_2.textChanged.connect(self.disable_if_empty)
        self.addport.ui.lineEdit_3.textChanged.connect(self.disable_if_empty)
        self.addport.ui.pushButton.setShortcut('Return')
        self.addport.show()

    def add_this_port(self):
        self.window.comboBox.currentTextChanged.disconnect()
        self.window.comboBox_2.currentTextChanged.disconnect()
        self.window.lineEdit.textChanged.disconnect()
        self.window.tableWidget.setRowCount(0)
        self.window.comboBox.clear()
        self.window.comboBox_2.clear()
        port = self.addport.ui.lineEdit.text()
        dest = self.addport.ui.lineEdit_2.text()
        code = self.addport.ui.lineEdit_3.text()
        self.window.tableWidget.setSortingEnabled(False)
        self.window.comboBox.addItem("Select destination name")
        self.window.comboBox_2.addItem("Select destination code")
        self.window.lineEdit.setText('')
        try:
            self.db.insert(port, code, dest)
            self.addport.ui.label.setText('Success!')
            self.addport.ui.label.setStyleSheet('color: green')
        except sqlite3.IntegrityError:
            self.addport.ui.label.setText('Not Inserted! Already Exists!')
            self.addport.ui.label.setStyleSheet('color: red')
        self.entries = self.db.get_all()
        self.last_entries = self.entries
        dropdown1 = set()
        dropdown2 = set()
        for entry in self.entries:
            dropdown1.add(entry[1])
            dropdown2.add(entry[2])
            row_position = self.window.tableWidget.rowCount()
            self.window.tableWidget.insertRow(row_position)
            self.window.tableWidget.setItem(row_position, 0, QTableWidgetItem(entry[0]))
            self.window.tableWidget.setItem(row_position, 1, QTableWidgetItem(entry[1]))
            self.window.tableWidget.setItem(row_position, 2, QTableWidgetItem(entry[2]))
        for d in sorted(dropdown1):
            self.window.comboBox.addItem(d)
        for d in sorted(dropdown2):
            self.window.comboBox_2.addItem(d)
        self.window.lineEdit.textChanged.connect(self.get_text)
        self.window.tableWidget.horizontalHeader().setSectionResizeMode(1)
        self.window.tableWidget.sortItems(0, Qt.AscendingOrder)
        self.window.comboBox.currentTextChanged.connect(self.get_text)
        self.window.comboBox_2.currentTextChanged.connect(self.get_text)
        self.window.pushButton_2.clicked.connect(self.add_dialog)
        self.window.tableWidget.setSortingEnabled(True)
        self.addport.ui.lineEdit.setText('')
        self.addport.ui.lineEdit.setFocus()
        pass

    def disable_if_empty(self):
        if self.addport.ui.lineEdit.text() == "" or self.addport.ui.lineEdit_2.text() == "" or self.addport.ui.lineEdit_3.text() == "":
            self.addport.ui.pushButton.setDisabled(True)
        else:
            self.addport.ui.pushButton.setDisabled(False)

    def get_text(self):
        self.window.comboBox.currentTextChanged.disconnect()
        self.window.comboBox_2.currentTextChanged.disconnect()
        self.window.lineEdit.textChanged.disconnect()
        combo = self.window.comboBox.currentText()
        combo2 = self.window.comboBox_2.currentText()
        box_text = self.window.lineEdit.text()
        self.window.comboBox.setCurrentIndex(-1)
        self.window.comboBox_2.setCurrentIndex(-1)
        self.window.tableWidget.setSortingEnabled(False)
        if box_text == '':
            isTextEmpty = True
        else:
            isTextEmpty = False
        if combo == "Select destination name":
            isDropdownDefault = True
        else:
            isDropdownDefault = False
        if combo2 == "Select destination code":
            isDropdown2Default = True
        else:
            isDropdown2Default = False
        if not isTextEmpty and isDropdownDefault and isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_specific_port(box_text)
        elif isTextEmpty and not isDropdownDefault and isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_specific_dest(combo)
        elif isTextEmpty and isDropdownDefault and not isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_specific_code(combo2)
        elif not isTextEmpty and not isDropdownDefault and isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_by_port_and_dest(box_text, combo)
        elif not isTextEmpty and isDropdownDefault and not isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_by_port_and_code(box_text, combo2)
        elif isTextEmpty and not isDropdownDefault and not isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_by_dest_and_code(combo, combo2)
        elif isTextEmpty and isDropdownDefault and isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.get_all()
        elif not isTextEmpty and not isDropdownDefault and not isDropdown2Default:
            self.window.tableWidget.setRowCount(0)
            self.window.comboBox.clear()
            self.window.comboBox_2.clear()
            self.window.comboBox.addItem("Select destination name")
            self.window.comboBox_2.addItem("Select destination code")
            self.entries = self.db.find_by_all(box_text, combo, combo2)
        dropdown1 = set()
        dropdown2 = set()
        for entry in self.entries:
            dropdown1.add(entry[1])
            dropdown2.add(entry[2])
            row_position = self.window.tableWidget.rowCount()
            self.window.tableWidget.insertRow(row_position)
            self.window.tableWidget.setItem(row_position, 0, QTableWidgetItem(entry[0]))
            self.window.tableWidget.setItem(row_position, 1, QTableWidgetItem(entry[1]))
            self.window.tableWidget.setItem(row_position, 2, QTableWidgetItem(entry[2]))
        for d in sorted(dropdown1):
            self.window.comboBox.addItem(d)
        for d in sorted(dropdown2):
            self.window.comboBox_2.addItem(d)
        self.window.comboBox.setCurrentText(combo)
        self.window.comboBox_2.setCurrentText(combo2)
        self.window.lineEdit.setText(box_text)
        self.window.lineEdit.textChanged.connect(self.get_text)
        self.window.tableWidget.horizontalHeader().setSectionResizeMode(1)
        self.window.tableWidget.sortItems(0, Qt.AscendingOrder)
        self.window.comboBox.currentTextChanged.connect(self.get_text)
        self.window.comboBox_2.currentTextChanged.connect(self.get_text)
        self.window.tableWidget.setSortingEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    ret = app.exec_()
    sys.exit(ret)
