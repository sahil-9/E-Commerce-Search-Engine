# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'win3.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!
from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd
from xlutils.copy import copy
import webbrowser

class Ui_Form2(object):
    
    temp1 = []
    name = " "
    price = " "
    link = " "
    senti = " "
    def dispdetails(self):
        fname = "E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test1.xls"
                                                                #opening file containing review 
        book2  = xlrd.open_workbook(fname)                           #opening xlsx file
        sheet2 = book2.sheet_by_index(0)
        wb2 = copy(book2)
        wsheet2 = wb2.get_sheet(0)
        self.name = str(sheet2.cell_value(0,0))
        self.price = str(sheet2.cell_value(0,1))
        self.link = str(sheet2.cell_value(0,2))
        self.senti = str(sheet2.cell_value(0,13))
        for j in range(3,13) : 
            self.temp1.append(str(sheet2.cell_value(0,j)))
        wb2.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test1.xls')
        
    def onReview(self): 
        k = 1   
        for i in range(len(self.temp1)):
            if self.temp1[i].strip(" ") == " ":
                break
            else:
                self.listWidget.addItem(str(k)+". "+self.temp1[i])
                k+=1

    def openlink(self):
        webbrowser.open(self.link)            
        
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1467, 965)
        Form.setWindowOpacity(1.0)
        Form.setStyleSheet("background-image: url(:/imgres/win3n.png);")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(0, 0, 1261, 101))
        self.label.setStyleSheet("font: 75 16pt \"Times New Roman\";\n"
"")
        self.dispdetails()
        self.label.setText("        Product Name ----> " +self.name + "\n        Price----->  " + self.price)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(10, 120, 131, 41))
        self.label_2.setStyleSheet("font: 75 16pt \"Times New Roman\";\n"
"text-decoration: underline;")
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(180, 120, 131, 41))
        self.pushButton.setStyleSheet("font: 75 14pt \"Times New Roman\";\n"
"text-decoration: underline;")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.openlink)
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(20, 380, 691, 571))
        self.label_3.setStyleSheet("background-image: url(:/imgres/white.png);")
        pyg = QtGui.QPixmap("Win3.png")
        self.label_3.setPixmap(pyg)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(0, 170, 651, 191))
        self.label_4.setStyleSheet("background-image: url(:/imgres/sentiment.png);\n"
"font: 75 italic 18pt \"Times New Roman\";")
        self.label_4.setText("               OverAll Sentiment : " + self.senti)
        self.label_4.setObjectName("label_4")
        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(990, 520, 461, 431))
        self.listWidget.setStyleSheet("background-image: url(:/imgres/review.png);\n"
"font: 75 14pt \"Times New Roman\";\n"
"background-image: url(:/imgres/review.jpg);")
        self.listWidget.setProperty("isWrapping", False)
        self.listWidget.setWordWrap(True)
        self.listWidget.setObjectName("listWidget")
        #self.onReview()
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(1130, 420, 191, 71))
        self.pushButton_2.setStyleSheet("font: 75 14pt \"Times New Roman\";\n"
"text-decoration: underline;")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.onReview)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label_2.setText(_translate("Form", "   To Order :"))
        self.pushButton.setText(_translate("Form", "Click Here"))
        self.pushButton_2.setText(_translate("Form", "Reviews"))

import win3n_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form2()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

