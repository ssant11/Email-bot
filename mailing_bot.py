#!/usr/bin/env python
# -*- coding: utf-8 -*-

# License:
# https://ms-python.gallerycdn.vsassets.io/extensions/ms-python/python/2021.2.582707922/1613774555212/Microsoft.VisualStudio.Services.Content.License


#hashtags above should not be removed if we want to use utf-8 characters
import pandas as pd
import smtplib
import codecs
from email.message import EmailMessage
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from validate_email import validate_email
import xlrd
import os
import sys

#Global variables
FEMALE = 'Mrs'
MALE = 'Mr'


class InvalidSexError(Exception): #error definition
    '''
    Class of error regarding invalid sex in Excel file.
    '''
    pass


class Ui_MainWindow(object):
    '''
    Class regarding the UI. Contains attributes:
    :param PATH: path of Excel file
    :type PATH: str

    :param content_fem: content of message directed towards females
    :type content_fem: str

    :param content_mal: content of message directed towards males
    :type content_mal: str

    :param signature: signature attached at the end of every email
    :type signature: str
    '''
    def __init__(self):
        '''
        Setting initial attribute values.
        '''
        self.PATH = ''

        self.content_fem = self.message_content(FEMALE)[0] #method call reading the email text
        self.content_mal = self.message_content(MALE)[0]
        self.signature = self.message_content(FEMALE)[1]


    def setupUi(self, MainWindow):
        '''
        Setting up the UI. Creating labels, buttons and editable text fields in MainWindow.
        '''
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 478)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.SEND_EMAIL = QtWidgets.QPushButton(self.centralwidget)
        self.SEND_EMAIL.setGeometry(QtCore.QRect(510, 350, 151, 41))
        self.SEND_EMAIL.setObjectName("SEND_EMAIL") #setting the name of an object
        self.SEND_EMAIL.clicked.connect(self.send_m) #setting the connection of the click with the method

        self.EMAIL_ADDRESS = QtWidgets.QLineEdit(self.centralwidget)
        self.EMAIL_ADDRESS.setGeometry(QtCore.QRect(100, 90, 181, 31))
        self.EMAIL_ADDRESS.setObjectName("EMAIL_ADDRESS")

        self.EMAIL_PASSWORD = QtWidgets.QLineEdit(self.centralwidget)
        self.EMAIL_PASSWORD.setGeometry(QtCore.QRect(340, 90, 181, 31))
        self.EMAIL_PASSWORD.setObjectName("EMAIL_PASSWORD")
        self.EMAIL_PASSWORD.setEchoMode(QtWidgets.QLineEdit.Password)

        self.EMAIL_LABEL = QtWidgets.QLabel(self.centralwidget)
        self.EMAIL_LABEL.setGeometry(QtCore.QRect(100, 50, 191, 31))
        self.EMAIL_LABEL.setObjectName("EMAIL_LABEL")

        self.PASSWORD_LABEL = QtWidgets.QLabel(self.centralwidget)
        self.PASSWORD_LABEL.setGeometry(QtCore.QRect(340, 50, 191, 31))
        self.PASSWORD_LABEL.setObjectName("PASSWORD_LABEL")

        self.EXCEL_PATH = QtWidgets.QLabel(self.centralwidget)
        self.EXCEL_PATH.setGeometry(QtCore.QRect(40, 210, 541, 31))
        self.EXCEL_PATH.setObjectName("EXCEL_PATH")
        self.EXCEL_PATH.setStyleSheet("background-color: white; border-width: '1px'; border-color: rgb(121, 121, 121); border-style: solid")

        self.BROWSE_EXCEL = QtWidgets.QPushButton(self.centralwidget)
        self.BROWSE_EXCEL.setGeometry(QtCore.QRect(600, 200, 151, 41))
        self.BROWSE_EXCEL.setObjectName("BROWSE_EXCEL")
        self.BROWSE_EXCEL.clicked.connect(self.browse_e)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def browse_e(self):
        '''
        Extracting file path from file explorer. Displaying the result on self.EXCEL_PATH label.
        '''
        self.PATH = QFileDialog.getOpenFileName()[0] #it returns tuple so we take the first element so the path to the file
        if self.PATH:
            self.EXCEL_PATH.setText(self.PATH) #set the label text as the path

    def check_emails(self, dataframe): #method for checking the emails correctness
        '''
        Verifying emails in terms of format (text@text.domain).
        '''
        for email in dataframe['Email']:
            is_valid = validate_email(str(email)) #checking if the mail has the correct form - no if it really exists!
            if not is_valid:
                return False
        return True

    def message_content(self, sex): #method that returns the email content based on gender
        '''
        Extracting email content from .txt files, with UTF-8 encoding. Returns tuple with content and signature.
        '''
        this_folder = os.path.dirname(os.path.abspath(__file__)) #capturing the location of this file .py
        my_file_f = os.path.join(this_folder, 'message_f.txt') #appending the filename to the folder location
        my_file_m = os.path.join(this_folder, 'message_m.txt')
        my_signature = os.path.join(this_folder, 'signature.txt')
        if str(sex) == FEMALE:
            with codecs.open(my_file_f, encoding='utf-8', mode='r') as entryfile: #we include unicode characters in the text
                content = entryfile.read() #the txt file is closed after exiting from with
        else:
            with codecs.open(my_file_m, encoding='utf-8', mode='r') as entryfile:
                content = entryfile.read()
        with codecs.open(my_signature, encoding='utf-8', mode='r') as entryfile:
            signature = entryfile.read()
        return (content, signature)

    def choose_content(self, sex, content_f, content_m): #selects the message content based on gender
        '''
        Returns correct message content having a given sex.
        '''
        if str(sex) == FEMALE:
            return content_f
        elif str(sex) == MALE:
            return content_m
        else:
            raise InvalidSexError

    def send_m(self): #sending e-mails
        '''
        Sending emails. Verifies all the content in regard of errors. If none detected, sends all the prepared emails.
        '''
        file_checker = False #control variable for errors related to an excel file
        try:
            df = pd.read_excel(self.PATH)
        except FileNotFoundError:
            file_checker = True
            self.popup_box("Error", "Path missing.")
        except xlrd.biffh.XLRDError:
            file_checker = True
            self.popup_box("Error", "Invalid file extension.")
        if not file_checker:
            EMAIL_ADDRESS = self.EMAIL_ADDRESS.text() #assigning the entered email address to a variable
            EMAIL_PASSWORD = self.EMAIL_PASSWORD.text() #assigning the entered password to a variable
            try:
                if len(df) == 0:
                    self.popup_box("Error", "Excel is empty.")
                else:
                    errchecker = False #bool variable for error checking to not send send confirmation when there is an error
                    message_list = [] #we save messages locally to be sent in bulk rather than sending them until an error occurs
                    for i in range(len(df)):
                        msg = EmailMessage()
                        msg['Subject'] = u'Customer Satisfaction Survey' #to be able to use utf-8 characters
                        msg['From'] = EMAIL_ADDRESS
                        try:
                            msg['To'] = df['Email'][i]
                        except KeyError:
                            errchecker = True
                            self.popup_box("Error", "Invalid column name(s). Check the Excel file.")
                            break
                        try:
                            msg.set_content(self.choose_content(df['Sex'][i], self.content_fem, self.content_mal) + '\n' + df['Link'][i] + '\n\n' + self.signature)
                        except TypeError:
                            errchecker = True
                            self.popup_box("Error", "Invalid link. Check the file for missing lines or links.")
                            break
                        except KeyError:
                            errchecker = True
                            self.popup_box("Error", "Invalid column name(s). Check the Excel file.")
                            break
                        except InvalidSexError:
                            errchecker = True
                            self.popup_box("Error", "Invalid sex. Check the Excel file.")
                            break
                        if not self.check_emails(df):
                            errchecker = True
                            self.popup_box("Error", "Invalid recipient address. Check the Excel file for missing or invalid email addresses.")
                            break
                        message_list.append(msg) #adding a message to the end of the list
                    if not errchecker: #sending emails after checking the correctness of excel (no duplicating emails to customers)
                        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp: #logging into the account for sending e-mails
                            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                            for messg in message_list:
                                smtp.send_message(messg) #sending messages from the list
                        self.popup_box("Success", "Email(s) sent.")

            except (smtplib.SMTPAuthenticationError, TypeError, UnicodeEncodeError):
                self.popup_box("Error", "Invalid email or password.")


    def popup_box(self, title, message):  #creating a pop up
        '''
        Pop-up box with designated title and error message. Used to display errors or to confirm that the emails have been sent.
        '''
        self.popup = QtWidgets.QMessageBox()
        self.popup.setWindowTitle(title)
        self.popup.setText(message)
        self.popup.exec_()

    def retranslateUi(self, MainWindow):
        '''
        Inner QtDesigner method.
        '''
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.SEND_EMAIL.setText(_translate("MainWindow", "Send"))
        self.EMAIL_LABEL.setText(_translate("MainWindow", "Email"))
        self.PASSWORD_LABEL.setText(_translate("MainWindow", "Password"))
        self.BROWSE_EXCEL.setText(_translate("MainWindow", "Browse"))



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())