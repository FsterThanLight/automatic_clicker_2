# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'update.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Update_UI(object):
    def setupUi(self, Update_UI):
        Update_UI.setObjectName("Update_UI")
        Update_UI.resize(283, 152)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        Update_UI.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/按钮图标/窗体/res/更新.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Update_UI.setWindowIcon(icon)
        self.verticalLayout = QtWidgets.QVBoxLayout(Update_UI)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.label = QtWidgets.QLabel(Update_UI)
        self.label.setMaximumSize(QtCore.QSize(60, 60))
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap(":/按钮图标/窗体/res/图标.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.label_2 = QtWidgets.QLabel(Update_UI)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.progressBar = QtWidgets.QProgressBar(Update_UI)
        self.progressBar.setStyleSheet("QProgressBar\n"
"{\n"
"    border: 1px solid #666666;\n"
"    text-align: center;\n"
"    color: #000;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QProgressBar::chunk\n"
"{\n"
"    background-color: red; /* 将黄色改为红色 */\n"
"    width: 5px;\n"
"    margin: 0.5px;\n"
"}\n"
"")
        self.progressBar.setProperty("value", 24)
        self.progressBar.setAlignment(QtCore.Qt.AlignCenter)
        self.progressBar.setOrientation(QtCore.Qt.Horizontal)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.progressBar)

        self.retranslateUi(Update_UI)
        QtCore.QMetaObject.connectSlotsByName(Update_UI)

    def retranslateUi(self, Update_UI):
        _translate = QtCore.QCoreApplication.translate
        Update_UI.setWindowTitle(_translate("Update_UI", "软件更新"))
        self.label_2.setText(_translate("Update_UI", "信息"))
import images_rc
