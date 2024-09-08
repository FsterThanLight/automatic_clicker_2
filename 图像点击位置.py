import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QPainter, QPen
from PyQt5.QtWidgets import QApplication, QLabel, QDialog

from 窗体.clickposition import Ui_ClickPosition

image_path_ = r'C:\Users\FS\Desktop\Clicker_test\qLcehE1sth.png'
image_path_2 = r'C:\Users\FS\Desktop\Clicker_test\T5pQ785JMV.png'


class MyLabel(QLabel):
    def __init__(self, parent, image_path=None, position='(0,0)'):
        super(MyLabel, self).__init__(parent)
        if image_path:
            self.pixmap = QPixmap(image_path)
        else:
            self.pixmap = QPixmap(100, 100)
            self.pixmap.fill(Qt.white)

        self.crosshair_color = Qt.red  # 十字框颜色
        self.crosshair_thickness = 2  # 十字框粗细
        # 设置十字框初始位置为图像中心
        if position != '(0,0)' and position != '(随机,随机)':
            position_ = eval(position)
            self.crosshair_position = (
                self.pixmap.width() // 2 + position_[0],
                self.pixmap.height() // 2 + position_[1],
            )
            self.parent().label_4.setText(str(position_[0]))
            self.parent().label_5.setText(str(position_[1]))
        else:
            self.crosshair_position = (
                self.pixmap.width() // 2,
                self.pixmap.height() // 2,
            )
        # 设置图像信息（宽度、高度）
        self.set_image_info()

    def set_image_info(self):
        img_width = self.pixmap.width()
        img_height = self.pixmap.height()
        self.parent().label_8.setText(str(img_width))
        self.parent().label_9.setText(str(img_height))

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        # 获取图像和窗口的尺寸
        img_width = self.pixmap.width()
        img_height = self.pixmap.height()
        win_width = self.width()
        win_height = self.height()

        # 计算图像的偏移量
        offset_x = (win_width - img_width) // 2
        offset_y = (win_height - img_height) // 2

        # 绘制图像，居中显示
        painter.drawPixmap(offset_x, offset_y, self.pixmap)

        # 获取十字框的位置
        x, y = self.crosshair_position
        painter.setPen(QPen(self.crosshair_color, self.crosshair_thickness, Qt.SolidLine))

        # 画水平线
        painter.drawLine(0, y + offset_y, self.width(), y + offset_y)
        # 画垂直线
        painter.drawLine(x + offset_x, 0, x + offset_x, self.height())

    def mousePressEvent(self, event):
        self.crosshair_position = (
            event.x() - (self.width() - self.pixmap.width()) // 2,
            event.y() - (self.height() - self.pixmap.height()) // 2,
        )
        self.update()

    def mouseMoveEvent(self, event):
        self.crosshair_position = (
            event.x() - (self.width() - self.pixmap.width()) // 2,
            event.y() - (self.height() - self.pixmap.height()) // 2,
        )
        self.update()
        # 计算十字框在图像中的位置，中心为0,0
        x = self.crosshair_position[0] - self.pixmap.width() // 2
        y = self.crosshair_position[1] - self.pixmap.height() // 2
        self.parent().label_4.setText(str(x))
        self.parent().label_5.setText(str(y))


class ClickPosition(QDialog, Ui_ClickPosition):
    def __init__(self, parent=None, image_path=None, position='(0,0)'):
        super().__init__(parent)
        self.setupUi(self)
        self.image_path = image_path
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )  # 隐藏帮助按钮
        # 创建自定义的MyLabel实例，传入图像路径
        self.label = MyLabel(self, self.image_path, position)
        if position == '(随机,随机)':
            self.checkBox.setChecked(True)
            self.label_4.setText('随机')
            self.label_5.setText('随机')
        # 在原label的位置插入自定义的label
        self.horizontalLayout.insertWidget(0, self.label)
        # 调整比例
        self.horizontalLayout.setStretch(0, 3)
        self.horizontalLayout.setStretch(1, 1)
        # 保存数据
        self.pushButton.clicked.connect(self.save_position)
        # 是否随机点击
        self.checkBox.stateChanged.connect(self.random_click)

    def save_position(self):
        try:
            tabWidget_title = self.parent().tabWidget.tabText(self.parent().tabWidget.currentIndex())
            if tabWidget_title == '图像点击':
                self.parent().label_176.setText(f'({self.label_4.text()},{self.label_5.text()})')
                self.close()
        except Exception as e:
            print(f'保存数据出现错误: {e}')

    def random_click(self):
        if self.checkBox.isChecked():
            self.label_4.setText('随机')
            self.label_5.setText('随机')
        else:
            self.label_4.setText('0')
            self.label_5.setText('0')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainwindow = ClickPosition(image_path=image_path_2, position='(50,50)')
    mainwindow.show()
    sys.exit(app.exec_())
