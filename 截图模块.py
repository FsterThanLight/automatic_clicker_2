import tkinter as tk

import pyautogui
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


class ScreenCapture:
    def __init__(self):
        self.rect = None  # 截图的矩形区域
        # 鼠标左键按下的位置
        self.x_1 = 0
        self.y_1 = 0
        # 鼠标左键抬起的位置
        self.x_3 = 0
        self.y_3 = 0
        self.pic = None
        self.region = None

    def screenshot_area(self):
        """截取屏幕矩形区域"""
        root = tk.Tk()
        root.attributes('-fullscreen', True)
        root.attributes('-alpha', 0.3)
        root.configure(bg='grey')

        # 鼠标左键按下事件
        def on_press(event):
            self.x_1, self.y_1 = event.x, event.y
            # print('鼠标开始点击位置为：', self.x_1, self.y_1)
            self.rect = canvas.create_rectangle(self.x_1, self.y_1, 0, 0, outline='red', width=2, fill='black')

        def on_drag(event):
            x_2, y_2 = event.x, event.y
            canvas.coords(self.rect, self.x_1, self.y_1, x_2, y_2)
            # 将矩形区域内的颜色设置为透明
            root.wm_attributes("-transparentcolor", "black")

        def on_release(event):
            self.x_3, self.y_3 = event.x, event.y
            # print('鼠标释放位置为：', self.x_3, self.y_3)
            canvas.delete(self.rect)
            canvas.destroy()
            root.destroy()

        # 创建屏幕遮罩
        canvas = tk.Canvas(root, bg="grey", cursor='cross')
        canvas.pack(fill=tk.BOTH, expand=True)
        # 绑定鼠标事件
        canvas.bind('<ButtonPress-1>', on_press)
        canvas.bind('<B1-Motion>', on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        # 开始事件循环
        root.mainloop()
        self.region = (self.x_1, self.y_1, self.x_3 - self.x_1, self.y_3 - self.y_1)

    def screenshot_region(self):
        """截取屏幕区域"""
        self.pic = pyautogui.screenshot(region=self.region)


if __name__ == '__main__':
    screen_capture = ScreenCapture()
    screen_capture.screenshot_area()
    screen_capture.screenshot_region()
