# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'win2.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!
from PyQt5 import QtCore, QtGui, QtWidgets
from PIL import Image
from win3 import Ui_Form2
import xlrd
import xlwt
from xlutils.copy import copy
class Ui_Form(object):
    
    prodlist = []
    
    def prod_list(self):
        fname = "E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls"
                                                                #opening file containing review 
        book  = xlrd.open_workbook(fname)                           #opening xlsx file
        sheet = book.sheet_by_index(0)
        wb = copy(book)
        wsheet = wb.get_sheet(0)
        
        for i in range(1,26):
            #for j in range(0,i):
            self.prodlist.append(str(i)+". "+str(sheet.cell_value(i,0)))
        
        k = 1
        for j in range(len(self.prodlist)):        
            self.listWidget.addItem(self.prodlist[j])
            k+=1
        wb.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls')
        
    def onGO(self):
        curr_prod = str(self.listWidget.currentItem().text())
        fname = "E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls"
                                                                #opening file containing review 
        book1  = xlrd.open_workbook(fname)                           #opening xlsx file
        sheet1 = book1.sheet_by_index(0)
        wb1 = copy(book1)
        wsheet1 = wb1.get_sheet(0)
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('A Test Sheet 1')
        
        name = " "
        price = " "
        link = " "
        temp = " "
        senti = " "
        for i in range(len(self.prodlist)):
            if self.prodlist[i] == curr_prod :
                ind = i
        
        index1 = ind+1 
        name = str(sheet1.cell_value(index1,0))
        price = str(sheet1.cell_value(index1,1))
        link = str(sheet1.cell_value(index1,2))
        senti = str(sheet1.cell_value(index1,17))
        worksheet.write(0,0,name)
        worksheet.write(0,1,price)
        worksheet.write(0,2,link)
        for j in range(3,13) : 
            temp = str(sheet1.cell_value(index1,j))
            worksheet.write(0,j,temp)
        worksheet.write(0,13,senti)
        wb1.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls')
        workbook.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test1.xls')
        
        image_name = str("product" + str(index1) + ".png")
        pieIMG =Image.open(image_name)    
        pieIMG.save("E:\Eclipse neon\Eclipse python\Final SDL\src\Win3.png")
        
        self.SW1 = QtWidgets.QMainWindow()
        self.ui = Ui_Form2()
        self.ui.setupUi(self.SW1)
        self.SW1.show()
                
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1314, 865)
        Form.setStyleSheet("background-image: url(:/imgres/win2new1.png);")
        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(50, 60, 561, 801))
        self.listWidget.setStyleSheet("background-image: url(:/imgres/qlist.png);\n"
"text-decoration: underline;\n"
"font: 75 14pt \"Times New Roman\";")
        self.listWidget.setProperty("isWrapping", False)
        self.listWidget.setWordWrap(True)
        self.listWidget.setObjectName("listWidget")
        self.prod_list()
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(870, 140, 221, 71))
        self.pushButton.setStyleSheet("font: 75 16pt \"Times New Roman\";")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.onGO)
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "win2"))
        self.pushButton.setText(_translate("Form", "GO"))

#import win1_rc
#import win2_rc
import win2n_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

