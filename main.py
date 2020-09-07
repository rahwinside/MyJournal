import sys

from datetime import datetime

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication
from openpyxl import workbook, load_workbook

START_TIME = datetime.now()
STOP_TIME = datetime.now()


class Ui_load(object):
    def setupUi(self, load):
        load.setObjectName("load")
        load.resize(775, 494)
        self.gridLayout = QtWidgets.QGridLayout(load)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setSizeConstraint(QtWidgets.QLayout.SetMinimumSize)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.startButton = QtWidgets.QPushButton(load)
        self.startButton.setObjectName("startButton")
        self.horizontalLayout.addWidget(self.startButton)
        self.stopButton = QtWidgets.QPushButton(load)
        self.stopButton.setObjectName("stopButton")
        self.horizontalLayout.addWidget(self.stopButton)
        self.gridLayout.addLayout(self.horizontalLayout, 2, 0, 1, 1)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(load)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.logText = QtWidgets.QTextEdit(load)
        self.logText.setObjectName("logText")
        self.verticalLayout.addWidget(self.logText)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_3 = QtWidgets.QLabel(load)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.statusLabel = QtWidgets.QLabel(load)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.statusLabel.sizePolicy().hasHeightForWidth())
        self.statusLabel.setSizePolicy(sizePolicy)
        self.statusLabel.setObjectName("statusLabel")
        self.horizontalLayout_2.addWidget(self.statusLabel)
        self.gridLayout.addLayout(self.horizontalLayout_2, 3, 0, 1, 1)

        self.retranslateUi(load)
        QtCore.QMetaObject.connectSlotsByName(load)

    def retranslateUi(self, load):
        _translate = QtCore.QCoreApplication.translate
        load.setWindowTitle(_translate("load", "MyJournal"))
        self.startButton.setText(_translate("load", "Start Timer"))
        self.stopButton.setText(_translate("load", "Stop Timer"))
        self.label.setText(_translate("load", "Worklog:"))
        self.label_3.setText(_translate("load", "Status:"))
        self.statusLabel.setText(_translate("load", "Starting application..."))
        self.startButton.setEnabled(True)
        self.stopButton.setEnabled(False)


class UX(QDialog, Ui_load):
    def __init__(self):
        QDialog.__init__(self)
        Ui_load.__init__(self)
        self.setupUi(self)
        self.startButton.clicked.connect(self.startTimerClicked)
        self.stopButton.clicked.connect(self.stopTimerClicked)
        self.statusLabel.setText("Ready")

    def startTimerClicked(self):
        global START_TIME

        self.statusLabel.setText("Toggling")
        self.startButton.setEnabled(False)
        self.stopButton.setEnabled(True)
        START_TIME = datetime.now()
        self.statusLabel.setText("Timer started")
        print(START_TIME)

    def stopTimerClicked(self):
        global START_TIME, STOP_TIME

        self.statusLabel.setText("Opening journal")
        workbook_name = 'journal.xlsx'
        wb = load_workbook(workbook_name)
        page = wb.active

        self.statusLabel.setText("Toggling")
        self.startButton.setEnabled(True)
        self.stopButton.setEnabled(False)
        STOP_TIME = datetime.now()
        self.statusLabel.setText("Timer stopped")

        log = self.logText.toPlainText()

        self.statusLabel.setText("Writing to journal")
        current_date = str(START_TIME).split(" ")[0]
        page.append([current_date, START_TIME, STOP_TIME, (STOP_TIME - START_TIME), log])
        wb.save(filename=workbook_name)
        self.statusLabel.setText("Written 1 row to journal")
        self.logText.setText("")

        print(STOP_TIME)
        print("Stopped")


app = QApplication(sys.argv)
ux = UX()
ux.show()
app.exec_()
