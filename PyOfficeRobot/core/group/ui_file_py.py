# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'ui_file.ui'
##
## Created by: Qt User Interface Compiler version 6.4.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QMetaObject, QRect,
                            QSize)
from PySide6.QtGui import (QFont)
from PySide6.QtWidgets import (QLabel, QListView, QPushButton)


class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(500, 600)
        Form.setMinimumSize(QSize(500, 600))
        Form.setMaximumSize(QSize(500, 600))
        self.listView = QListView(Form)
        self.listView.setObjectName(u"listView")
        self.listView.setGeometry(QRect(10, 50, 341, 541))
        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(20, 10, 321, 31))
        font = QFont()
        font.setPointSize(20)
        self.label.setFont(font)
        self.pushButton_load = QPushButton(Form)
        self.pushButton_load.setObjectName(u"pushButton_load")
        self.pushButton_load.setGeometry(QRect(360, 50, 131, 141))
        font1 = QFont()
        font1.setPointSize(30)
        self.pushButton_load.setFont(font1)
        self.pushButton_start = QPushButton(Form)
        self.pushButton_start.setObjectName(u"pushButton_start")
        self.pushButton_start.setGeometry(QRect(360, 410, 131, 141))
        self.pushButton_start.setFont(font1)
        self.pushButton_close = QPushButton(Form)
        self.pushButton_close.setObjectName(u"pushButton_close")
        self.pushButton_close.setGeometry(QRect(360, 230, 131, 141))
        self.pushButton_close.setFont(font1)

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)

    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.label.setText(QCoreApplication.translate("Form", u"\u7fa4\u53d1\u5217\u8868 : ", None))
        self.pushButton_load.setText(QCoreApplication.translate("Form", u"\u52a0\u8f7d", None))
        self.pushButton_start.setText(QCoreApplication.translate("Form", u"\u7fa4\u53d1", None))
        self.pushButton_close.setText(QCoreApplication.translate("Form", u"\u6682\u505c", None))
    # retranslateUi
