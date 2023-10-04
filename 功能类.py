import time

import pyautogui
from PyQt5.QtWidgets import QApplication, QMessageBox

from main_work import sqlitedb, close_database


def timer(func):
    def func_wrapper(*args, **kwargs):
        from time import time
        time_start = time()
        result = func(*args, **kwargs)
        time_end = time()
        time_spend = time_end - time_start
        print('%s cost time: %.3f s' % (func.__name__, time_spend))
        return result

    return func_wrapper


def get_setting_data_from_db() -> tuple:
    """从数据库中获取设置参数
    :return: 设置参数"""
    cursor, conn = sqlitedb()
    cursor.execute('select * from 设置')
    list_setting_data = cursor.fetchall()
    close_database(cursor, conn)
    # 使用字典来存储设置参数
    setting_dict = {i[0]: i[1] for i in list_setting_data}
    return (setting_dict.get('持续时间'),
            setting_dict.get('时间间隔'),
            setting_dict.get('图像匹配精度'),
            setting_dict.get('暂停时间'))


class ImageClick:
    """图像点击"""

    def __init__(self, main_window, ins_dic):
        # 设置参数
        (self.duration,
         self.interval,
         self.confidence,
         self.time_sleep) = get_setting_data_from_db()
        # 主窗口
        self.main_window = main_window
        # 指令字典
        self.ins_dic = ins_dic

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数
        :return: 指令参数列表，重复次数"""
        # 读取图像名称
        img = self.ins_dic.get('图像路径')
        # 取重复次数
        re_try = self.ins_dic.get('重复次数')
        # 是否跳过参数
        skip = self.ins_dic.get('参数2')
        click_map = {
            '左键单击': [1, 'left', img, skip],
            '左键双击': [2, 'left', img, skip],
            '右键单击': [1, 'right', img, skip],
            '右键双击': [2, 'right', img, skip]
        }
        list_ins = click_map.get(self.ins_dic.get('参数1（键鼠指令）'))
        # 返回重复次数，点击次数，左键右键，图片名称，是否跳过
        return re_try, list_ins[0], list_ins[1], list_ins[2], list_ins[3]

    def start_execute(self, number):
        """开始执行鼠标点击事件
        :param number: 主窗口显示的循环次数"""
        # 解析指令字典
        reTry, click_times, lOrR, img, skip = self.parsing_ins_dic()
        # 执行图像点击
        if reTry == 1:
            self.execute_click(click_times, lOrR, img, skip, number)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.execute_click(click_times, lOrR, img, skip, number)
                i += 1
                time.sleep(self.time_sleep)

    def execute_click(self, click_times, lOrR, img, skip, number):
        """执行鼠标点击事件"""
        # 4个参数：鼠标点击时间，按钮类型（左键右键中键），图片名称，重复次数
        repeat = True
        number_1 = 1

        def image_match_click(skip_, start_time_=None):
            nonlocal repeat, number_1
            if location is not None:
                # 参数：位置X，位置Y，点击次数，时间间隔，持续时间，按键
                self.main_window.plainTextEdit.appendPlainText('找到匹配图片' + str(number))
                pyautogui.click(location.x, location.y,
                                clicks=click_times,
                                interval=self.interval,
                                duration=self.duration,
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
                location = pyautogui.locateCenterOnScreen(
                    image=img, confidence=self.confidence)
                image_match_click(skip)
            else:
                while repeat:
                    location = pyautogui.locateCenterOnScreen(
                        image=img, confidence=self.confidence)
                    image_match_click(skip, start_time)
        except OSError:
            QMessageBox.critical(self.main_window,
                                 '错误', '文件下未找到.png图像文件，请检查文件是否存在！')


class Coordinate_Click:
    def __init__(self, main_window, ins_dic):
        # 设置参数
        (self.duration,
         self.interval,
         self.confidence,
         self.time_sleep) = get_setting_data_from_db()
        # 主窗口
        self.main_window = main_window
        # 指令字典
        self.ins_dic = ins_dic

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        re_try = self.ins_dic.get('重复次数')
        # 取x,y坐标的值
        x_ = int(self.ins_dic.get('参数2').split('-')[0])
        y_ = int(self.ins_dic.get('参数2').split('-')[1])
        z_ = int(self.ins_dic.get('参数2').split('-')[1])
        self.main_window.plainTextEdit.appendPlainText('x_,y坐标：' + str(x_) + ',' + str(y_))
        click_map = {
            '左键单击': [1, 'left', x_, y_],
            '左键双击': [2, 'left', x_, y_],
            '右键单击': [1, 'right', x_, y_],
            '右键双击': [2, 'right', x_, y_],
            '左键（自定义次数）': [z_, 'left', x_, y_]
        }
        list_ins = click_map.get(self.ins_dic.get('参数1（键鼠指令）'))
        # 返回重复次数，点击次数，左键右键，x坐标，y坐标
        return re_try, list_ins[0], list_ins[1], list_ins[2], list_ins[3]

    def start_execute(self, number):
        """开始执行鼠标点击事件
        :param number: 主窗口显示的循环次数"""
        # 获取参数
        reTry, click_times, lOrR, x__, y__ = self.parsing_ins_dic()
        # 执行坐标点击
        if reTry == 1:
            pyautogui.click(x__, y__, click_times,
                            interval=self.interval,
                            duration=self.duration,
                            button=lOrR)
            self.main_window.plainTextEdit.appendPlainText(
                '执行坐标%s:%s点击' % (x__, y__) + str(number))
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                pyautogui.click(x__, y__, click_times,
                                interval=self.interval,
                                duration=self.duration,
                                button=lOrR)
                self.main_window.plainTextEdit.appendPlainText(
                    '执行坐标%s:%s点击' % (x__, y__) + str(number))
                i += 1
                time.sleep(self.time_sleep)


if __name__ == '__main__':
    pass
    x, y, z, w = get_setting_data_from_db()
    print(x, y, z, w)
