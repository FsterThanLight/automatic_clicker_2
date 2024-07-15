import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel,
                             QPushButton, QColorDialog, QVBoxLayout)


class DemoColorDialog(QWidget):
    def __init__(self, parent=None):
        super(DemoColorDialog, self).__init__(parent)

        # 设置窗口标题
        self.setWindowTitle('实战PyQt5: QColorDialog Demo!')
        # 设置窗口大小
        self.resize(360, 240)

        self.initUi()

    def initUi(self):
        vLayout = QVBoxLayout(self)
        vLayout.addSpacing(10)

        btnTest = QPushButton('调整颜色', self)
        btnTest.clicked.connect(self.onSetFont)

        self.label_text = QLabel('实战PyQt5: \n测试QColorDialog')
        self.label_text.setAlignment(Qt.AlignCenter)
        self.label_text.setFont(QtGui.QFont(self.font().family(), 16))

        vLayout.addWidget(btnTest)
        vLayout.addWidget(self.label_text)

        self.setLayout(vLayout)

    def onSetFont(self):
        col = QColorDialog.getColor()
        pal = self.label_text.palette()
        pal.setColor(QPalette.WindowText, col)
        print('颜色:', col)
        print('颜色:', pal)
        self.label_text.setPalette(pal)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DemoColorDialog()
    window.show()
    sys.exit(app.exec())
