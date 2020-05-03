import sys
import ntpath
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5 import QtWidgets, QtGui
from exdifftool import run

def color_table_row(table, rowIndex, color):
    for j in range(table.columnCount()):
        if table.item(rowIndex, j):
            table.item(rowIndex, j).setBackground(color)

class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Excel Diff Tool'
        self.left = 0
        self.top = 0
        self.width = 1650
        self.height = 800
        self.left_filename = ""
        self.right_filename = ""
        self.initUI()
        
    def update(self, left_csv, right_csv, diffs):
    
        self.fill_table(self.tableWidgetL, left_csv)
        self.fill_table(self.tableWidgetR, right_csv)
        self.color_tables(diffs)
        
    def fill_table(self, tableWidget, csv):
        j = 0
        for row in csv:
            cells = row.split(",")
            i = 0
            for cell in cells:
                tableWidget.setItem(j, i, QTableWidgetItem(cell))
                i += 1
            j += 1
        
        tableWidget.setRowCount(j)
        tableWidget.setColumnCount(i)
        header = tableWidget.horizontalHeader()       
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        
    def color_tables(self, diffs):
        for diff in diffs:
            op = diff[0].split(" ")[0]
            col1 = int(diff[0].split(" ")[1].split(":")[0])
            col2 = int(diff[0].split(" ")[1].split(":")[1])
            
            if op == "neq":
                color_table_row(self.tableWidgetL, col1, QtGui.QColor(255, 200, 200))
                color_table_row(self.tableWidgetR, col2, QtGui.QColor(255, 200, 200))
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
        scrollL.setWidgetResizable(True)
        scrollL.setWidget(self.tableWidgetL)
        layoutL.addWidget(scrollL)
        
        # right side
        self.loadButtonR = QPushButton("<click to select file 2>", self)
        self.loadButtonR.clicked.connect(self.openFileNameDialogRight)
        layoutR.addWidget(self.loadButtonR)
        scrollR.setWidgetResizable(True)
        scrollR.setWidget(self.tableWidgetR)
        layoutR.addWidget(scrollR)
        
        layoutMain.addLayout(layoutL)
        layoutMain.addLayout(layoutR)
        self.setLayout(layoutMain)

        # Show widget
        self.show()

    def createTableLeft(self):
       # Create table
        self.tableWidgetL = QTableWidget()
        self.tableWidgetL.move(0,0)
        self.tableWidgetL.setRowCount(10)
        self.tableWidgetL.setColumnCount(10)
        
    def createTableRight(self):
       # Create table
        self.tableWidgetR = QTableWidget()
        self.tableWidgetR.move(0,0)
        self.tableWidgetR.setRowCount(10)
        self.tableWidgetR.setColumnCount(10)
        
    def openFileNameDialogLeft(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Excel Files (*.xlsx)", options=options)
        if fileName:
            self.left_filename = fileName
            self.loadButtonL.setText(ntpath.basename(fileName))
            if self.right_filename:
                # do diff
                self.doDiff()
                
    def openFileNameDialogRight(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Excel Files (*.xlsx)", options=options)
        if fileName:
            self.right_filename = fileName
            self.loadButtonR.setText(ntpath.basename(fileName))
            if self.left_filename:
                # do diff
                self.doDiff()
                
    def doDiff(self):
        csv1, csv2, diffs = run(self.left_filename, self.right_filename)
        self.update(csv1, csv2, diffs)

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
