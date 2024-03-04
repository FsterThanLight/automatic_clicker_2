from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QPen, QColor
from PyQt5.QtWidgets import QWidget


class TransparentWindow(QWidget):
    """显示框选区域的窗口"""

    def __init__(self, pos):
        """pos(x,y, width, height)"""
        super().__init__()
        # 设置无边框窗口
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setWindowOpacity(0.5)  # 设置透明度
        self.setAttribute(Qt.WA_TranslucentBackground)  # 设置背景透明
        self.setGeometry(pos[0], pos[1], pos[2], pos[3])  # 设置窗口大小

    def paintEvent(self, event):
        # 绘制边框
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(QColor(255, 0, 0), 5, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        painter.drawRect(self.rect())


if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    window = TransparentWindow((100, 100, 400, 400))
    window.show()
    sys.exit(app.exec_())
