import sys
import ntpath
import logging
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5 import QtWidgets, QtGui
from exdifftool import run, get_sheet_names, get_row_and_cols_of_csv

def color_table_row(table, row, color):
    for j in range(table.columnCount()):
        if table.item(row, j):
            table.item(row, j).setBackground(color)

def color_table_item(table, row, col, color):
    if table.item(row, col):
        table.item(row, col).setBackground(color)
        
class App(QWidget):

    def __init__(self):
        super().__init__()
        logging.basicConfig(stream=sys.stderr, level=logging.INFO)
        self.title = 'Excel Diff Tool'
        self.left = 0
        self.top = 0
        self.width = 1650
        self.height = 800
        self.left_filename = ""
        self.right_filename = ""
        self.left_sheet = -1
        self.right_sheet = -1
        self.initUI()
        
    def update(self, left_csv, right_csv, diffs):
        self.fill_table(self.tableWidgetL, left_csv)
        self.fill_table(self.tableWidgetR, right_csv)
        self.color_tables(diffs)
        
    def fill_table(self, tableWidget, csv):
        n_rows, n_cols = get_row_and_cols_of_csv(csv)
        tableWidget.setRowCount(n_rows)
        tableWidget.setColumnCount(n_cols)
        
        j = 0
        for row in csv:
            cells = row.split(",")
            i = 0
            for cell in cells:
                tableWidget.setItem(j, i, QTableWidgetItem(cell))
                i += 1
            j += 1
        
        header = tableWidget.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        
    def color_tables(self, diffs):
        for diff in diffs:
            op = diff.split(" ")[0]
            col1 = int(diff.split(" ")[1].split(":")[0])
            col2 = int(diff.split(" ")[1].split(":")[1])
            
            if op == "neqd":
                color_table_item(self.tableWidgetL, col1, col2, QtGui.QColor(255, 200, 200))
                color_table_item(self.tableWidgetR, col1, col2, QtGui.QColor(255, 200, 200))
            elif op == "mv":
                color_table_row(self.tableWidgetL, col1, QtGui.QColor(255, 255, 200))
                color_table_row(self.tableWidgetR, col2, QtGui.QColor(255, 255, 200))
            elif op == "add":
                if col1 != -1:
                    color_table_row(self.tableWidgetL, col1, QtGui.QColor(200, 255, 200))
                elif col2 != -1:
                    color_table_row(self.tableWidgetR, col2, QtGui.QColor(200, 255, 200))
        
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        
        self.createTableLeft()
        self.createTableRight()

        layoutL = QVBoxLayout()
        layoutR = QVBoxLayout()
        layoutMain = QHBoxLayout()
        
        scrollL = QScrollArea()
        scrollR = QScrollArea()
        
        self.tableWidgetL.horizontalScrollBar().valueChanged.connect(self.tableWidgetR.horizontalScrollBar().setValue)
        self.tableWidgetR.horizontalScrollBar().valueChanged.connect(self.tableWidgetL.horizontalScrollBar().setValue)
        self.tableWidgetL.verticalScrollBar().valueChanged.connect(self.tableWidgetR.verticalScrollBar().setValue)
        self.tableWidgetR.verticalScrollBar().valueChanged.connect(self.tableWidgetL.verticalScrollBar().setValue)
        
        # left side
        self.loadButtonL = QPushButton("<click to select file 1>", self)
        self.loadButtonL.clicked.connect(self.openFileNameDialogLeft)
        layoutL.addWidget(self.loadButtonL)
        self.sheetCBoxL = QComboBox(self)
        self.sheetCBoxL.addItem("Sheet 1")
        self.sheetCBoxL.currentIndexChanged.connect(self.leftSheetSelected)
        layoutL.addWidget(self.sheetCBoxL)
        scrollL.setWidgetResizable(True)
        scrollL.setWidget(self.tableWidgetL)
        layoutL.addWidget(scrollL)
        
        
        # right side
        self.loadButtonR = QPushButton("<click to select file 2>", self)
        self.loadButtonR.clicked.connect(self.openFileNameDialogRight)
        layoutR.addWidget(self.loadButtonR)
        self.sheetCBoxR = QComboBox(self)
        self.sheetCBoxR.addItem("Sheet 1")
        self.sheetCBoxR.currentIndexChanged.connect(self.rightSheetSelected)
        layoutR.addWidget(self.sheetCBoxR)
        scrollR.setWidgetResizable(True)
        scrollR.setWidget(self.tableWidgetR)
        layoutR.addWidget(scrollR)
        
        # legend
        tbox = QLineEdit(self)
        tbox.setReadOnly(True)
        tbox.setText("equal")
        layoutL.addWidget(tbox)
        
        tbox = QLineEdit(self)
        tbox.setReadOnly(True)
        tbox.setText("changed")
        tbox.setStyleSheet("background-color: rgb(255, 200, 200);")
        layoutL.addWidget(tbox)
        
        tbox = QLineEdit(self)
        tbox.setReadOnly(True)
        tbox.setStyleSheet("background-color: rgb(255, 255, 200);")
        tbox.setText("moved")
        layoutR.addWidget(tbox)
        
        tbox = QLineEdit(self)
        tbox.setReadOnly(True)
        tbox.setText("added")
        tbox.setStyleSheet("background-color: rgb(200, 255, 200);")
        layoutR.addWidget(tbox)
        
        layoutMain.addLayout(layoutL)
        layoutMain.addLayout(layoutR)
        
        self.setLayout(layoutMain)

        # Show widget
        self.show()
        
    def rightSheetSelected(self):
        self.right_sheet = self.sheetCBoxR.currentIndex()
        self.doDiff()
        
    def leftSheetSelected(self):
        self.left_sheet = self.sheetCBoxL.currentIndex()
        self.doDiff()

    def createTableLeft(self):
       # Create table
        self.tableWidgetL = QTableWidget()
        self.tableWidgetL.move(0,0)
        self.tableWidgetL.setRowCount(50)
        self.tableWidgetL.setColumnCount(50)
        
    def createTableRight(self):
       # Create table
        self.tableWidgetR = QTableWidget()
        self.tableWidgetR.move(0,0)
        self.tableWidgetR.setRowCount(50)
        self.tableWidgetR.setColumnCount(50)
        
    def openFileNameDialogLeft(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Excel Files (*.xlsx)", options=options)
        if fileName:
            self.left_filename = fileName
            self.loadButtonL.setText(ntpath.basename(fileName))
            # set sheet names
            sheet_names = get_sheet_names(fileName)
            self.sheetCBoxL.clear()
            for s in sheet_names:
                self.sheetCBoxL.addItem(s)
                
    def openFileNameDialogRight(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Excel Files (*.xlsx)", options=options)
        if fileName:
            self.right_filename = fileName
            self.loadButtonR.setText(ntpath.basename(fileName))
            # set sheet names
            sheet_names = get_sheet_names(fileName)
            self.sheetCBoxR.clear()
            for s in sheet_names:
                self.sheetCBoxR.addItem(s)
                
    def doDiff(self):
        if self.left_filename and self.right_filename and self.left_sheet >= 0 and self.right_sheet >= 0:
            csv1, csv2, diffs = run(self.left_filename, self.right_filename, self.left_sheet, self.right_sheet)
            self.update(csv1, csv2, diffs)

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
