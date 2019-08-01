# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from win2 import Ui_Form

class Ui_window1(object):
    
    def onStart(self):
       
        prod_name = self.lineEdit.text()
        
       # Form = QtWidgets.QWidget()
        #Form.hide()
        
        import flipkart
        flipkart.search_prod(prod_name)
        import Shopclues
        Shopclues.searchprod(prod_name)
        import SentimentalAnalysis
        SentimentalAnalysis.sentimental_analysis()
        import pie_chart
        pie_chart.pie()
        self.SW = QtWidgets.QMainWindow()
        self.ui = Ui_Form()
        self.ui.setupUi(self.SW)
        self.SW.show()
         
    def setupUi(self, window1):
        window1.setObjectName("window1")
        window1.resize(1071, 722)
        window1.setStyleSheet("background-image: url(:/imgres/win1back.png);\n"
"background-repeat : no-repeat;")
        self.lineEdit = QtWidgets.QLineEdit(window1)
        self.lineEdit.setGeometry(QtCore.QRect(150, 460, 791, 41))
        self.lineEdit.setStyleSheet("background-image: url(:/imgres/txt.png);\n"
"font: 75 16pt \"Times New Roman\";")
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(window1)
        self.pushButton.setGeometry(QtCore.QRect(460, 530, 201, 51))
        self.pushButton.setStyleSheet("font: 75 14pt \"MS Shell Dlg 2\";\n"
"background-image: url(:/imgres/brow.png);\n"
"color:white;")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.onStart)
        self.label = QtWidgets.QLabel(window1)
        self.label.setGeometry(QtCore.QRect(270, 90, 521, 151))
        self.label.setStyleSheet("background-image: url(:/imgres/omo.jpg);")
        #self.label.setText("")
        self.label.setObjectName("label")

        self.retranslateUi(window1)
        QtCore.QMetaObject.connectSlotsByName(window1)

    def retranslateUi(self, window1):
        _translate = QtCore.QCoreApplication.translate
        window1.setWindowTitle(_translate("window1", "window1"))
        self.pushButton.setText(_translate("window1", "SEARCH"))

import win1_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window1 = QtWidgets.QWidget()
    ui = Ui_window1()
    ui.setupUi(window1)
    window1.show()
    sys.exit(app.exec_())
