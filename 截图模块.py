import io
import os
import random
import string
import tkinter as tk

import pyautogui
from PyQt5 import QtGui
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtWidgets import QDialog, QMessageBox

from ini控制 import extract_resource_folder_path
from 窗体.image_preview import Ui_Image


class ImagePreview(QDialog, Ui_Image):
    def __init__(self, im_bytes, im_b, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.im_bytes = im_bytes  # 图片的二进制数据
        self.im_b = im_b  # 图片的二进制数据
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.load_setting_data()  # 加载设置数据
        self.lineEdit.setValidator(QRegExpValidator(QRegExp("[a-zA-Z0-9]{16}"), self))
        # 按钮事件
        self.pushButton.clicked.connect(self.save_image)
        self.lineEdit.setText(f'{self.generate_random_alphanumeric(10)}')

    def load_setting_data(self):
        folder_path_list = extract_resource_folder_path()
        self.comboBox.clear()
        self.comboBox.addItems(folder_path_list)

    def preview_image(self):
        """预览图片"""
        self.label.setPixmap(QtGui.QPixmap.fromImage(QtGui.QImage.fromData(self.im_b)))

    @staticmethod
    def generate_random_alphanumeric(length):
        # 生成随机字母和数字的组合
        characters = string.ascii_letters + string.digits
        # 从字符集中随机选择字符，重复 length 次，并将结果连接成字符串
        return ''.join(random.choice(characters) for _ in range(length))

    def save_image(self):
        """保存图片"""
        folder_path = self.comboBox.currentText()  # 选择的文件夹路径
        if os.path.exists(folder_path):  # 如果文件路径存在
            file_name_ = self.lineEdit.text().strip()
            file_name = file_name_ if file_name_.endswith('.png') else file_name_ + '.png'
            # 拼接文件路径
            file_path = os.path.join(folder_path, file_name)
            # 保存图片
            with open(file_path, 'wb') as f:
                f.write(self.im_bytes.getvalue())
        else:
            QMessageBox.warning(self, '警告', '文件夹路径不存在，保存失败！')
        # 关闭窗口
        self.accept()


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
            self.rect = canvas.create_rectangle(
                self.x_1, self.y_1, 0, 0, outline='red', width=2, fill='black'
            )

        def on_drag(event):
            x_2, y_2 = event.x, event.y
            canvas.coords(self.rect, self.x_1, self.y_1, x_2, y_2)
            # 将矩形区域内的颜色设置为透明
            root.wm_attributes("-transparentcolor", "black")

        def on_release(event):
            self.x_3, self.y_3 = event.x, event.y
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
        # 检查位置是否异常，right是否小于left，bottom是否小于top
        if self.region[2] < 0:
            self.region = (self.region[0] + self.region[2], self.region[1], -self.region[2], self.region[3])
        if self.region[3] < 0:
            self.region = (self.region[0], self.region[1] + self.region[3], self.region[2], -self.region[3])

    def screenshot_region(self):
        """截取屏幕区域"""
        self.pic = pyautogui.screenshot(region=self.region)

    def show_preview(self):
        """显示截图预览"""
        # 保存截图到内存
        im_bytes = io.BytesIO()
        self.pic.save(im_bytes, format='png')
        im_b = im_bytes.getvalue()
        # 显示截图
        image_preview = ImagePreview(im_bytes, im_b)
        image_preview.preview_image()
        image_preview.exec_()
        # 释放内存
        del im_bytes
        del im_b


if __name__ == '__main__':
    screen_capture = ScreenCapture()
    screen_capture.screenshot_area()
    screen_capture.screenshot_region()
