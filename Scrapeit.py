# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'bot_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from main_scraper import scrape
from time import sleep
from threading import *
from PyQt5.QtWidgets import QMessageBox


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(400, 300)
        MainWindow.setMinimumSize(QtCore.QSize(400, 300))
        MainWindow.setMaximumSize(QtCore.QSize(400, 300))
        MainWindow.setBaseSize(QtCore.QSize(400, 400))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.start_button = QtWidgets.QPushButton(self.centralwidget)
        self.start_button.setGeometry(QtCore.QRect(130, 200, 121, 51))
        self.start_button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.start_button.setObjectName("start_button")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 60, 381, 61))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.text_label = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.text_label.setMaximumSize(QtCore.QSize(400, 300))
        font = QtGui.QFont()
        font.setPointSize(17)
        self.text_label.setFont(font)
        self.text_label.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.text_label.setAlignment(QtCore.Qt.AlignCenter)
        self.text_label.setObjectName("text_label")
        self.verticalLayout.addWidget(self.text_label)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "S&P Scraper"))
        self.start_button.setText(_translate("MainWindow", "START BOT"))
        self.text_label.setText(_translate("MainWindow", "CLICK TO START THE BOT "))
        self.start_button.clicked.connect(self.start_button_clicked)

    def start_button_clicked(self):
        self.text_label.setText("scraper started please wait !")
        print("start button clicked")
        self.start_button.setEnabled(False)
        self.text_label.setText("scraper Initialized")
        t1 = Thread(target=self.start)
        t1.start()

    def show_info_messagebox(self,title,message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        
        msg.setWindowTitle(title)
        msg.setText(message)
        retval = msg.exec_()

    def start(self):
        try:
            sc = scrape()
            for i in range(1, 12):
                self.text_label.setText(f"pages Scraped : {i}")
                url = f"https://markets.businessinsider.com/index/components/s&p_500?p={i}"
                sc.goto_url(url)
                sleep(4)
                sc.scrape_table()
        except Exception as e:
            # sc.exit()
            print(e)
        finally:
            self.text_label.setText("Scraping Done")
            sleep(1)
            self.text_label.setText('Now Scraping News')
            sc.scrape_news()
            self.text_label.setText("styling and making excel file")
            sc.convert_xls('output.csv')
            print("done")
            self.show_info_messagebox("Information","Done ! Close the application")
            sc.exit()
    
        



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
