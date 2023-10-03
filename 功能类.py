import time

import pyautogui
from PyQt5.QtWidgets import QApplication, QMessageBox


class ImageClick:
    """图像点击"""

    def __init__(self,
                 img,
                 click_times,
                 lOrR,
                 skip,
                 interval,
                 duration,
                 confidence,
                 main_window,
                 time_sleep):

        # 功能参数
        self.img = img  # 图片路径
        self.click_times = click_times  # 点击次数
        self.lOrR = lOrR  # 左键还是右键
        self.skip = skip  # 跳过的次数
        # 设置参数
        self.interval = interval  # 点击间隔
        self.duration = duration  # 持续时间
        self.confidence = confidence  # 图像匹配精度
        self.settings = time_sleep  # 暂停时间
        # 窗口参数
        self.main_window = main_window  # 主窗口

    def start_execute(self, reTry, number):
        """开始执行鼠标点击事件
        :param number: 主窗口显示的循环次数
        :param reTry: 重试次数"""
        if reTry == 1:
            self.execute_click(self.click_times, self.lOrR, self.img, self.skip, number)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.execute_click(self.click_times, self.lOrR, self.img, self.skip, number)
                i += 1
                time.sleep(self.settings.time_sleep)

    def execute_click(self, click_times, lOrR, img, skip, number):
        """执行鼠标点击事件
        :param number: 主窗口显示的循环次数
        :param click_times: 点击次数
        :param lOrR: 左键还是右键（left,right）
        :param img: 图片路径
        :param skip: 跳过的次数"""

        # 4个参数：鼠标点击时间，按钮类型（左键右键中键），图片名称，重复次数
        repeat = True
        number_1 = 1

        def image_match_click(skip_, start_time_=None):
            nonlocal repeat, number_1
            if location is not None:
                # 参数：位置X，位置Y，点击次数，时间间隔，持续时间，按键
                self.main_window.plainTextEdit.appendPlainText('找到匹配图片' + str(number))
                pyautogui.click(location.x, location.y,
                                clicks=click_times, interval=self.interval, duration=self.duration,
                                button=lOrR)
                repeat = False
            else:
                if skip_ != "自动略过":
                    # 计算如果时间差的秒数大于skip则退出
                    # 获取当前时间，计算时间差
                    end_time = time.time()
                    time_difference = end_time - start_time_
                    # 显示剩余等待时间
                    self.main_window.plainTextEdit.appendPlainText(
                        '未找到匹配图片' + str(number) + '正在重试第' + str(number_1) + '次')
                    self.main_window.plainTextEdit.appendPlainText(
                        '剩余等待' + str(round(int(skip_) - time_difference, 0)) + '秒')
                    number_1 += 1
                    QApplication.processEvents()
                    # 终止条件
                    if time_difference > int(skip_):
                        repeat = False
                        raise pyautogui.ImageNotFoundException
                    time.sleep(0.1)
                elif skip_ == "自动略过":
                    self.main_window.plainTextEdit.appendPlainText('未找到匹配图片' + str(number))

        try:
            start_time = time.time()
            if skip == "自动略过":
                location = pyautogui.locateCenterOnScreen(image=img, confidence=self.confidence)
                image_match_click(skip)
            else:
                while repeat:
                    location = pyautogui.locateCenterOnScreen(image=img, confidence=self.confidence)
                    image_match_click(skip, start_time)
        except OSError:
            QMessageBox.critical(self.main_window, '错误', '文件下未找到.png图像文件，请检查文件是否存在！')
