# -*- coding: utf-8 -*-
"""
Created on Wed May 12 13:55:49 2021

@author: Egon
"""
import sys
import locale
from datetime import datetime
import win32com.client
from sql import *
from PyQt5.QtCore import *
from PyQt5 import QtCore
from PyQt5.QtGui import *
from PyQt5 import QtGui, uic
from PyQt5.QtWidgets import *
from PyQt5.uic import *
from PyQt5 import QtPrintSupport
import csv
from os.path import exists
from os import mkdir
import os
from ErrorLogger import *
import traceback

from reportFile import reportObj
 
trans = QtCore.QTranslator()

class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)

        global helper
        global db_path
       
        try:  
            current_db_file = open(r"assets/db/current_db.txt", "r+")
            db_path = current_db_file.readline()
            current_db_file.close()
            
            if os.path.isfile(db_path): 
                helper = SqlHelper(db_path)
            else:
                helper = SqlHelper('kassenbuch.db')
                
                cuurent_db_file = open(r"assets/db/current_db.txt", "w")
                cuurent_db_file.write('.//kassenbuch.db')
                cuurent_db_file.close()
                db_path = '.\\kassenbuch.db'
            
            last_language= load_last_language()
            load_message(last_language)        
            change_language(last_language, 'kassenbuch.ui', self)
            self.set_date_time()
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))          

    def keyPressEvent(self, event):
           if event.key() == 16777220 or event.key() == 16777221: #Enter
               self.buchen()
          
    def refresh(self):
        self.handle_menu()
        self.handel_Buttons()        
        self.lineEdit.clear()
        self.fill_kassenbestand()
        self.fill_buchungen()
    def handel_Buttons(self):
        self.pushButton_3.clicked.connect(self.beenden)   #Beenden
        self.pushButton.clicked.connect(self.journal)     #Journal
        self.pushButton_2.clicked.connect(self.storno)    #Storno
        self.pushButton_4.clicked.connect(self.export_as_csv)
   
    def handle_menu(self):
        try:
            self.actionNew.triggered.connect(self.new_database)
            self.actionOpen.triggered.connect(self.open_database) 
            self.actionEnglish.triggered.connect(lambda: [load_message("en"), change_language("en", 'kassenbuch.ui', self),\
                                                          set_last_language("en")])
            self.actionDeutsch.triggered.connect(lambda: [load_message("de"), change_language("de", 'kassenbuch.ui', self),\
                                                          set_last_language("de")])
            self.actionEspanol.triggered.connect(lambda: [load_message("es"), change_language("es", 'kassenbuch.ui', self),\
                                                          set_last_language("es")])
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
            
    def set_date_time(self):
        self.dateEdit.setDateTime(QDateTime.currentDateTime())

    def fill_kassenbestand(self):
        try:
            summeEinnahmeData = helper.select("SELECT Sum(einnahme) FROM buchung;")
            if summeEinnahmeData[0][0] is None:
                summeEinnahme = 0
            else:
                summeEinnahme = int(summeEinnahmeData[0][0])
            
            summeAusgabeData = helper.select("SELECT Sum(ausgabe) FROM buchung;")
            if summeAusgabeData[0][0] is None:
                summeAusgabe = 0
            else:
                summeAusgabe = summeAusgabeData[0][0]          
            kassenbestand = summeEinnahme - summeAusgabe
            self.lineEdit.setText(str(kassenbestand))        #Aktueller Kassenbestand
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
        
    def fill_buchungen(self):
        try:
            self.tableWidget.setRowCount(0)
            data = helper.select("Select buchungsid,datum,einnahme, ausgabe,buchungstext,belegnr,mwst from buchung") 
            for row , form in enumerate(data):
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)  
                for column , item in enumerate(form):
                    # if column ==1:
                    #     self.tableWidget.setItem(row , column , QTableWidgetItem(str(item.strftime("%d.%m.%Y"))))
                    # else:
                    self.tableWidget.setItem(row , column , QTableWidgetItem(str(item)))
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
            
    def beenden(self):
        self.close()
    def journal(self):
        self.window2 = datumOrtwaehlen()
        self.window2.show()
        
    def storno(self):
        try:
            if self.tableWidget.item(self.tableWidget.currentRow(),0):
                ei = self.tableWidget.item(self.tableWidget.currentRow(),2).text() #Einnahme
                au = self.tableWidget.item(self.tableWidget.currentRow(),3).text() #Ausgabe
                textStorno = messageText[11][0]
                buchungsid = self.tableWidget.item(self.tableWidget.currentRow(),0).text()
                if (not ei =='0'):
                    ei =0
                    updates = (ei,textStorno,buchungsid)
                    helper.edit("UPDATE buchung SET einnahme=?, buchungstext=? WHERE buchungsid=?",updates)
                if (not au =='0'):
                    au=0
                    updates = (au,textStorno,buchungsid)
                    helper.edit("UPDATE buchung SET ausgabe=?, buchungstext=? WHERE buchungsid=?",updates)
                QMessageBox.warning(self, 'Info', messageText[11][0], QMessageBox.Ok) 
            else:
                QMessageBox.warning(self, 'Info', messageText[16][0], QMessageBox.Ok) 
                
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))          
        
        self.fill_kassenbestand()    
        self.fill_buchungen()
    def buchen(self):
        try: 
            da = self.dateEdit.date().toPyDate()
            ei = self.lineEdit_6.text() #Einnahme
            au = self.lineEdit_5.text() #Ausgabe
            bu = self.lineEdit_4.text() #Buchungstext
            be = self.lineEdit_3.text() #BelegNr
            ms = self.comboBox.currentText() #MwSt 
            if (not ei and not au) or (ei and au):
                QMessageBox.warning(self, 'Error' , messageText[5][0], QMessageBox.Ok)
                return
            if ei:
                if (not ei.isdigit()):
                    QMessageBox.warning(self, 'Error', messageText[6][0], QMessageBox.Ok)
                    self.lineEdit_6.setFocus()
                    return
                buchungssatz = (da,ei,bu,be,ms)
                helper.insert("INSERT INTO buchung (datum,einnahme,buchungstext,belegnr,mwst) Values(?,?,?,?,?)",buchungssatz)
            if au:
                if (not au.isdigit()):
                    QMessageBox.warning(self, 'Error', messageText[7][0], QMessageBox.Ok)
                    self.lineEdit_5.setFocus()
                    return            
                buchung = (da,au,bu,be,ms)
                helper.insert("INSERT INTO buchung (datum,ausgabe,buchungstext,belegnr,mwst) Values(?,?,?,?,?)",buchung)
            self.lineEdit_6.clear() #Einnahme
            self.lineEdit_5.clear() #Ausgabe
            self.lineEdit_4.clear() #Buchungstext
            self.lineEdit_3.clear() #BelegNr
            self.fill_buchungen()
            self.fill_kassenbestand()
            self.statusBar().showMessage(messageText[1][0],5000)        
            self.lineEdit_6.setFocus() 
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))   
    def new_database(self):
        try:
            target = QFileDialog.getSaveFileName(self, messageText[13][0],messageText[13][0])
            target = (target[0])
            kassenbuch_db_file = r"assets/db/db_copy/kassenbuch.accdb"
            if target:
                import shutil
    
                original = kassenbuch_db_file
                target = target + ".accdb"
    
                shutil.copyfile(original, target)
                self.open_database(target)
                self.refresh()
                self.statusBar().showMessage(messageText[2][0],5000)  
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))             

    def open_database(self, db_file):
        global helper
        global db_path
        
        try:
            if not db_file:
                options = QFileDialog.Options()
                options |= QFileDialog.DontUseNativeDialog
                db_file, _ = QFileDialog.getOpenFileName(self, messageText[14][0], "",
                                                      "Access File (*.accdb)", options=options)
            if db_file:
                try:
                    helper = SqlHelper(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                        r'DBQ=' + db_file)
                    cuurent_db_file = open(r"assets/db/current_db.txt", "w")
                    cuurent_db_file.write(db_file)
                    cuurent_db_file.close()
                    db_path = db_file
                    self.statusBar().showMessage(messageText[3][0],5000)  
                except Exception as e:
                    ErrorLogger.WriteError('Line 208: ' + str(e))
                    QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
                    
            self.lineEdit.clear()
            self.fill_kassenbestand()
            self.fill_buchungen() 
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))               
    
    def export_as_csv(self):
        folder = "Export"
        if not exists(folder):
            mkdir(folder)
        file = "export" + str(datetime.now().strftime("%Y_%m_%d_%H_%M")) + ".csv"
        try:
            with open(folder + "/" + file, "w", newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                len = self.tableWidget.rowCount()
                if not len:
                    QMessageBox.warning(self, 'Error', messageText[8][0], QMessageBox.Ok)
                    return None
                for row in range(len):
                    if row == 0:
                        header = []
                    csv_row = []
                    for column in range(self.tableWidget.columnCount()):
                        if row == 0:
                            _item = self.tableWidget.horizontalHeaderItem(column)
                            if _item:
                                item = self.tableWidget.horizontalHeaderItem(column).text()
                                header.append(item)
                            else:
                                header.append("")

                        _item = self.tableWidget.item(row, column)
                        if _item:
                            item = self.tableWidget.item(row, column).text()
                            csv_row.append(item)
                        else:
                            csv_row.append("")
                    if row == 0:
                        csv_writer.writerow(header)
                    csv_writer.writerow(csv_row)
                    self.statusBar().showMessage(messageText[4][0],5000)        
                    
        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
        
class datumOrtwaehlen(QWidget):
    def __init__(self):
        QWidget.__init__(self)
        
        last_language = load_last_language()
        if  last_language == 'de':
            uic.loadUi(r"datumOrtwaehlen.ui", self)
        elif last_language == 'en':
            change_language('enDOW', 'datumOrtwaehlen.ui', self)
        elif last_language == 'es':
            change_language('esDOW', 'datumOrtwaehlen.ui', self)
        self.handle_button_do()
        self.set_date()
        
    def refresh(self):pass    
    def handle_button_do(self):
        self.toolButton.clicked.connect(self.select_folder)
        self.pushButton.clicked.connect(self.handlePrint)
        self.dir_path_to_save_files = None

    def set_date(self):
        self.dateEdit.setDate(QDate.currentDate())
        self.dateEdit_2.setDate(QDate.currentDate())

    def select_folder(self):
        the_dir_path = QFileDialog.getExistingDirectory()
        self.dir_path_to_save_files = the_dir_path
        self.lineEdit.setText(the_dir_path)
        
    def handlePrint(self):
        try:
            if not self.dir_path_to_save_files:
                QMessageBox.warning(self, 'Error', messageText[10][0], QMessageBox.Ok)
                return
     
            StartDate = self.dateEdit.date().toPyDate()
            EndDate = self.dateEdit_2.date().toPyDate()
          
            data = helper.select(f"select * from buchung where datum between '{StartDate}' and '{EndDate}'")
            language = load_last_language()
            # create a Report 
            report = reportObj(language, data, StartDate, EndDate)
            report.generate_journal_report("Journal.pdf")

        except Exception as e:
            ErrorLogger.WriteError(traceback.format_exc())
            QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e)) 
#@Language
def load_last_language():
    try:
        language_file = open(r"last_language.txt", "r+")
        last_language = language_file.readline()
        language_file.close()  
        return last_language  
    except Exception as e:
        ErrorLogger.WriteError(traceback.format_exc())
        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))           
def set_last_language(language):
    try:
        language_file = open(r"last_language.txt", "w")
        language_file.write(language)
        language_file.close() 
    except Exception as e:
        ErrorLogger.WriteError(traceback.format_exc())
        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))        
def load_message(language):
    try:
        global messageText
        messageText = helper.select(f"Select text from messagetext where language like '{language}' order by messageid")             
    except Exception as e:
        ErrorLogger.WriteError(traceback.format_exc())
        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))
def change_language(language, callingWindow, callingClass):
    try:
        if language:
            trans.load(language)
            QtWidgets.QApplication.instance().installTranslator(trans)
            uic.loadUi(callingWindow, callingClass)
            callingClass.refresh()
        else:
            QtWidgets.QApplication.instance().removeTranslator(trans) 
    except Exception as e:
        ErrorLogger.WriteError(traceback.format_exc())
        QtWidgets.QMessageBox.critical(None, 'Exception raised', format(e))            
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
    
if __name__ == '__main__':
    main()   