import io
import os
import random
import re
import sys
import time
import tkinter as tk
from datetime import datetime
from tkinter import ttk

import keyboard
import mouse
import openpyxl
import pyautogui
import pymsgbox
import pyperclip
import pyttsx4
import win32con
import win32gui
import win32process
import winsound
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QPen, QColor
from PyQt5.QtWidgets import QApplication, QWidget
from aip import AipOcr
from dateutil.parser import parse

from 数据库操作 import get_setting_data_from_db, get_str_now_time, get_variable_info, set_variable_value, \
    line_number_increment, get_ocr_info
from 网页操作 import WebOption

# sys.coinit_flags = 2  # STA
# from pywinauto import Application
# from pywinauto.findwindows import ElementNotFoundError

# dic_ = {
#                     'ID': elem_[0],
#                     '图像路径': elem_[1],
#                     '指令类型': elem_[2],
#                     '参数1（键鼠指令）': elem_[3],
#                     '参数2': elem_[4],
#                     '参数3': elem_[5],
#                     '参数4': elem_[6],
#                     '重复次数': elem_[7],
#                     '异常处理': elem_[8]
#                 }

DRIVER = None  # 浏览器驱动


def exit_main_work():
    sys.exit()


def close_browser():
    """关闭浏览器"""
    global DRIVER
    web_option = WebOption()
    web_option.driver = DRIVER
    web_option.close_browser()


def sub_variable(text: str):
    """将text中的变量替换为变量值"""
    new_text = text
    if ('☾' in text) and ('☽' in text):
        variable_dic = get_variable_info('dict')
        for key, value in variable_dic.items():
            new_text = new_text.replace(f'☾{key}☽', str(value))
    return new_text


class TransparentWindow(QWidget):
    """显示框选区域的窗口"""

    def __init__(self):
        """pos(x,y, width, height)"""
        super().__init__()
        # 设置无边框窗口
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setWindowOpacity(0.5)  # 设置透明度
        self.setAttribute(Qt.WA_TranslucentBackground)  # 设置背景透明
        # self.setGeometry(pos[0], pos[1], pos[2], pos[3])  # 设置窗口大小

    def paintEvent(self, event):
        # 绘制边框
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(QColor(255, 0, 0), 5, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        painter.drawRect(self.rect())


class OutputMessage:
    """输出信息，测试时输出到文本框，非测试时输出到主窗口"""

    def __init__(self, command_thread, navigation):
        # 输出的窗口
        self.command_thread = command_thread
        self.navigation = navigation

    def out_mes(self, message: str, is_test: bool = False):
        """输出信息,测试时输出到文本框，非测试时输出到主窗口"""
        if not is_test:
            self.command_thread.show_message(message)
        elif is_test:
            self.navigation.textBrowser.append(
                f'{get_str_now_time()}\t{message}'
            )
        QApplication.processEvents()


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


class ImageClick:
    """图像点击"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        setting_data_dic = get_setting_data_from_db(
            '持续时间', '时间间隔', '图像匹配精度', '暂停时间'
        )
        self.duration = float(setting_data_dic.get('持续时间'))
        self.interval = float(setting_data_dic.get('时间间隔'))
        self.confidence = float(setting_data_dic.get('图像匹配精度'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数
        :return: 指令参数列表，重复次数"""
        # 读取图像名称
        img = self.ins_dic.get('图像路径')
        # 取重复次数
        re_try = self.ins_dic.get('重复次数')
        skip = self.ins_dic.get('参数2')  # 是否跳过参数
        gray_recognition = eval(self.ins_dic.get('参数3'))  # 是否灰度识别
        area_identification = eval(self.ins_dic.get('参数4'))  # 是否区域识别
        if area_identification == (0, 0, 0, 0):
            area_identification = None  # 如果没有区域识别则设置为None
        click_map = {
            '左键单击': [1, 'left', img, skip],
            '左键双击': [2, 'left', img, skip],
            '右键单击': [1, 'right', img, skip],
            '右键双击': [2, 'right', img, skip],
            '仅移动鼠标': [0, 'left', img, skip]
        }
        list_ins = click_map.get(self.ins_dic.get('参数1（键鼠指令）'))
        # 返回重复次数，点击次数，左键右键，图片名称，是否跳过
        return re_try, gray_recognition, area_identification, list_ins[0], list_ins[1], list_ins[2], list_ins[3]

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 解析指令字典
        reTry, gray_rec, area, click_times, lOrR, img, skip = self.parsing_ins_dic()
        # 执行图像点击
        if reTry == 1:
            self.execute_click(click_times, gray_rec, lOrR, img, skip, area)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.execute_click(click_times, gray_rec, lOrR, img, skip, area)
                i += 1
                time.sleep(self.time_sleep)

    def execute_click(self, click_times, gray_rec, lOrR, img, skip, area=None):
        """执行鼠标点击事件
        :param click_times: 点击次数
        :param gray_rec: 是否灰度识别
        :param lOrR: 左键右键
        :param img: 图像名称
        :param skip: 是否跳过
        :param area: 是否区域识别"""

        # 4个参数：鼠标点击时间，按钮类型（左键右键中键），图片名称，重复次数
        def image_match_click(location, spend_time):
            if location is not None:
                if not self.is_test:
                    # 参数：位置X，位置Y，点击次数，时间间隔，持续时间，按键
                    self.out_mes.out_mes(f'已找到匹配图片，耗时{spend_time}毫秒。', self.is_test)
                    pyautogui.click(location.x, location.y,
                                    clicks=click_times,
                                    interval=self.interval,
                                    duration=self.duration,
                                    button=lOrR)
                elif self.is_test:
                    self.out_mes.out_mes(f'已找到匹配图片，耗时{spend_time}毫秒。', self.is_test)
                    # 移动鼠标到图片位置
                    pyautogui.moveTo(location.x, location.y, duration=0.2)

        try:
            min_search_time = 1 if skip == "自动略过" else float(skip)
            # 显示信息
            self.out_mes.out_mes(f'正在查找匹配图像...', self.is_test)
            QApplication.processEvents()
            # 记录开始时间
            start_time = time.time()
            location_ = pyautogui.locateCenterOnScreen(
                image=img,
                confidence=self.confidence,
                minSearchTime=min_search_time,
                grayscale=gray_rec,
                region=area
            )
            if location_:  # 如果找到图像
                spend_time_ = int((time.time() - start_time) * 1000)  # 计算耗时
                image_match_click(location_, spend_time_)
            elif not location_:  # 如果未找到图像
                self.out_mes.out_mes('未找到匹配图像', self.is_test)
            QApplication.processEvents()
        except OSError:
            self.out_mes.out_mes('文件下未找到png图像，请检查文件是否存在！', self.is_test)
            raise FileNotFoundError


class CoordinateClick:
    """坐标点击"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        setting_data_dic = get_setting_data_from_db(
            '持续时间', '时间间隔', '图像匹配精度', '暂停时间'
        )
        self.duration = float(setting_data_dic.get('持续时间'))
        self.interval = float(setting_data_dic.get('时间间隔'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        re_try = self.ins_dic.get('重复次数')
        # 取x,y坐标的值
        x_ = int(self.ins_dic.get('参数2').split('-')[0])
        y_ = int(self.ins_dic.get('参数2').split('-')[1])
        z_ = int(self.ins_dic.get('参数2').split('-')[2])
        click_map = {
            '左键单击': [1, 'left', x_, y_],
            '左键双击': [2, 'left', x_, y_],
            '右键单击': [1, 'right', x_, y_],
            '右键双击': [2, 'right', x_, y_],
            '左键（自定义次数）': [z_, 'left', x_, y_],
            '仅移动鼠标': [0, 'left', x_, y_]
        }
        list_ins = click_map.get(self.ins_dic.get('参数1（键鼠指令）'))
        # 返回重复次数，点击次数，左键右键，x坐标，y坐标
        return re_try, list_ins[0], list_ins[1], list_ins[2], list_ins[3]

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 获取参数
        reTry, click_times, lOrR, x__, y__ = self.parsing_ins_dic()
        # 执行坐标点击
        if reTry == 1:
            self.coor_click(click_times, lOrR, x__, y__)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.coor_click(click_times, lOrR, x__, y__)
                i += 1
                time.sleep(self.time_sleep)

    def coor_click(self, click_times, lOrR, x__, y__):
        pyautogui.click(x=x__, y=y__,
                        clicks=click_times,
                        interval=self.interval,
                        duration=self.duration,
                        button=lOrR
                        )
        if click_times == 0:
            self.out_mes.out_mes(f'移动鼠标到{x__}:{y__}', self.is_test)
        else:
            self.out_mes.out_mes(f'执行坐标{x__}:{y__}点击', self.is_test)


class TimeWaiting:
    """等待"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 执行线程
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def start_execute(self):
        """从指令字典中解析出指令参数"""
        wait_type = self.ins_dic.get('参数1（键鼠指令）')
        if wait_type == '时间等待':
            wait_time = int(self.ins_dic.get('参数2'))
            self.out_mes.out_mes('等待时长%d秒' % wait_time, self.is_test)
            self.stop_time(wait_time)
        elif wait_type == '定时等待':
            target_time, interval_time = self.ins_dic.get('参数2').split('+')
            # 检查目标时间是否大于当前时间
            if parse(target_time) > datetime.now():
                self.wait_to_time(target_time, interval_time)
        elif wait_type == '随机等待':
            min_time, max_time = self.ins_dic.get('参数2').split('-')
            wait_time = random.randint(int(min_time), int(max_time))
            self.out_mes.out_mes('随机等待时长%d秒' % wait_time, self.is_test)
            self.stop_time(wait_time)

    def wait_to_time(self, target_time, interval):
        """检查时间，指定时间则执行操作
        :param target_time: 目标时间
        :param interval: 时间间隔"""
        sleep_time = int(interval) / 1000
        show_times = 1  # 显示时间的间隔

        while True:
            now = datetime.now()
            if show_times == 1:
                self.out_mes.out_mes('当前为：%s' % now.strftime('%Y/%m/%d %H:%M:%S'), self.is_test)
                self.out_mes.out_mes('等待至：%s' % target_time, self.is_test)
                show_times = sleep_time
            if now >= parse(target_time):
                self.out_mes.out_mes('退出等待', self.is_test)
                break
            # 时间暂停
            time.sleep(sleep_time)
            show_times += sleep_time

    def stop_time(self, seconds):
        """暂停时间"""
        for i in range(seconds):
            # 显示剩下等待时间
            self.out_mes.out_mes('等待中...剩余%d秒' % (seconds - i), self.is_test)
            time.sleep(1)


class ImageWaiting:
    """图片等待"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def wait_to_image(self, image, wait_instruction_type, timeout_period):
        """执行图片等待"""
        if wait_instruction_type == '等待到指定图像出现':
            self.out_mes.out_mes('正在等待指定图像出现中...', self.is_test)
            QApplication.processEvents()
            location = pyautogui.locateCenterOnScreen(
                image=image,
                confidence=0.8,
                minSearchTime=timeout_period
            )
            if location:
                self.out_mes.out_mes('目标图像已经出现，等待结束', self.is_test)
                QApplication.processEvents()
        elif wait_instruction_type == '等待到指定图像消失':
            vanish = True
            while vanish:
                try:
                    pyautogui.locateCenterOnScreen(
                        image=image,
                        confidence=0.8,
                        minSearchTime=1
                    )
                except pyautogui.ImageNotFoundException:
                    self.out_mes.out_mes('目标图像已经消失，等待结束', self.is_test)
                    QApplication.processEvents()
                    vanish = False
                else:
                    time.sleep(0.5)

    def start_execute(self):
        """执行图片等待"""
        image_path = self.ins_dic.get('图像路径')
        wait_instruction_type = self.ins_dic.get('参数1（键鼠指令）')
        timeout_period = float(self.ins_dic.get('参数2'))
        self.wait_to_image(image_path, wait_instruction_type, timeout_period)


class RollerSlide:
    """滑动鼠标滚轮"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self, type_):
        """解析指令字典"""
        if type_ == '滚轮滑动':
            scroll_direction = str(self.ins_dic.get('参数2').split(',')[0])
            scroll_distance_ = int(self.ins_dic.get('参数2').split(',')[1])
            scroll_distance = scroll_distance_ if scroll_direction == '↑' else -scroll_distance_
            return scroll_direction, scroll_distance
        elif type_ == '随机滚轮滑动':
            min_distance = int(self.ins_dic.get('参数2').split(',')[0])
            max_distance = int(self.ins_dic.get('参数2').split(',')[1])
            scroll_direction = random.choice(['↑', '↓'])
            scroll_distance_ = random.randint(min_distance, max_distance)
            scroll_distance = scroll_distance_ if scroll_direction == '↑' else -scroll_distance_
            return scroll_direction, scroll_distance

    def start_execute(self):
        """执行重复次数"""
        type_ = self.ins_dic.get('参数1（键鼠指令）')
        re_try = self.ins_dic.get('重复次数')
        scroll_direction, scroll_distance = self.parsing_ins_dic(type_)
        # 执行滚轮滑动
        if re_try == 1:
            self.wheel_slip(scroll_direction, scroll_distance, type_)
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.wheel_slip(scroll_direction, scroll_distance, type_)
                i += 1
                time.sleep(self.time_sleep)

    def wheel_slip(self, scroll_direction, scroll_distance, type_):
        """滚轮滑动事件"""
        pyautogui.scroll(scroll_distance)
        self.out_mes.out_mes(f'{type_}{scroll_direction}{scroll_distance}距离', self.is_test)


class TextInput:
    """输入文本"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        setting_data_dic = get_setting_data_from_db('时间间隔', '暂停时间')
        self.interval = float(setting_data_dic.get('时间间隔'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def start_execute(self):
        """解析指令字典"""
        input_value = sub_variable(self.ins_dic.get('参数1（键鼠指令）'))
        special_control_judgment = eval(self.ins_dic.get('参数2'))
        # 执行文本输入
        self.text_input(input_value, special_control_judgment)

    def text_input(self, input_value, special_control_judgment):
        """文本输入事件
        :param input_value: 输入的文本
        :param special_control_judgment: 是否为特殊控件"""
        if not special_control_judgment:
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(self.time_sleep)
            self.out_mes.out_mes('执行文本输入：%s' % input_value, self.is_test)
        elif special_control_judgment:
            pyautogui.typewrite(input_value, interval=self.interval)
            self.out_mes.out_mes('执行特殊控件的文本输入%s' % input_value, self.is_test)
            time.sleep(self.time_sleep)


class MoveMouse:
    """移动鼠标"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        setting_data_dic = get_setting_data_from_db('持续时间', '暂停时间')
        self.duration = float(setting_data_dic.get('持续时间'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self, type_):
        """解析指令字典"""
        if type_ == '移动鼠标':
            direction = self.ins_dic.get('参数2').split(',')[0]
            distance = self.ins_dic.get('参数2').split(',')[1]
            return direction, distance
        elif type_ == '随机移动鼠标':
            random_type = self.ins_dic.get('参数2')
            return random_type

    def start_execute(self):
        """执行重复次数"""
        re_try = self.ins_dic.get('重复次数')
        type_ = self.ins_dic.get('参数1（键鼠指令）')
        # 执行滚轮滑动
        if re_try == 1:
            self.mouse_move_fun(type_)  # 执行鼠标移动
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.mouse_move_fun(type_)  # 执行鼠标移动
                i += 1
                time.sleep(self.time_sleep)

    def mouse_move_fun(self, type_: str) -> None:
        """执行鼠标移动
        :param type_: 鼠标移动类型"""
        if type_ == '移动鼠标':
            direction, distance = self.parsing_ins_dic(type_)
            self.mouse_moves(direction, distance)
        elif type_ == '随机移动鼠标':
            random_type = self.parsing_ins_dic(type_)
            if random_type == '类型1':
                self.mouse_moves_random_1()
            elif random_type == '类型2':
                self.mouse_moves_random_2()

    def mouse_moves(self, direction, distance):
        """鼠标移动事件"""
        # 相对于当前位置移动鼠标
        directions = {'↑': (0, -1), '↓': (0, 1), '←': (-1, 0), '→': (1, 0)}
        if direction in directions:
            x, y = directions.get(direction)
            pyautogui.moveRel(x * int(distance), y * int(distance), duration=self.duration)
        self.out_mes.out_mes('移动鼠标%s%s像素距离' % (direction, distance), self.is_test)

    def mouse_moves_random_1(self):
        """鼠标移动事件"""
        screen_width, screen_height = pyautogui.size()
        # 随机生成坐标
        x = random.randint(0, screen_width)
        y = random.randint(0, screen_height)
        # 随机生成时间
        duration_ran = random.uniform(0.1, 0.9)
        try:
            pyautogui.moveTo(x, y, duration=duration_ran)
            self.out_mes.out_mes('随机移动鼠标', self.is_test)
        except pyautogui.FailSafeException:
            pass

    def mouse_moves_random_2(self):
        """鼠标移动事件"""
        directions = {'↑': (0, -1), '↓': (0, 1), '←': (-1, 0), '→': (1, 0)}
        direction = random.choice(list(directions.keys()))
        if direction in directions:
            x, y = directions.get(direction)
            distance = random.randint(1, 500)
            duration_ran = random.uniform(0.1, 0.9)
            try:
                pyautogui.moveRel(x * distance, y * distance, duration=duration_ran)
                self.out_mes.out_mes('随机移动鼠标', self.is_test)
            except pyautogui.FailSafeException:
                pass


class PressKeyboard:
    """模拟按下键盘"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def parsing_ins_dic(self):
        """解析指令字典"""
        re_try = self.ins_dic.get('重复次数')
        key = self.ins_dic.get('参数1（键鼠指令）')
        return re_try, key

    def start_execute(self):
        """执行重复次数"""
        re_try, key = self.parsing_ins_dic()
        # 执行滚轮滑动
        if re_try == 1:
            self.press_keyboard(key)
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.press_keyboard(key)
                i += 1
                time.sleep(self.time_sleep)

    def press_keyboard(self, key):
        """鼠标移动事件
        :param key: 按键列表"""
        keyboard.press_and_release(key)
        self.out_mes.out_mes('按下按键%s' % key, self.is_test)


class MiddleActivation:
    """鼠标中键激活"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def start_execute(self):
        """执行重复次数"""
        command_type = self.ins_dic.get('参数1（键鼠指令）')
        click_count = int(self.ins_dic.get('参数2'))
        re_try = self.ins_dic.get('重复次数')
        # 执行滚轮滑动
        if re_try == 1:
            self.middle_mouse_button(command_type, click_count)
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.middle_mouse_button(command_type, click_count)
                i += 1
                time.sleep(self.time_sleep)

    def middle_mouse_button(self, command_type, click_times):
        """中键点击事件"""
        self.out_mes.out_mes('等待按下鼠标中键中...按下F11键退出', self.is_test)
        QApplication.processEvents()
        mouse.wait(button='middle')
        try:
            if command_type == "模拟点击":
                self.simulated_mouse_click(click_times, '左键')
                self.out_mes.out_mes(f'执行鼠标点击{click_times}次', self.is_test)
            elif command_type == "自定义":
                pass
        except OSError:
            # 弹出提示框。提示检查鼠标是否连接
            self.out_mes.out_mes('连接失败，请检查鼠标是否连接正确。', self.is_test)

    @staticmethod
    def simulated_mouse_click(click_times, lOrR):
        """模拟鼠标点击
        :param click_times: 点击次数
        :param lOrR: (左键、右键)"""
        button = 'left' if lOrR == '左键' else 'right'
        for i in range(click_times):
            mouse.click(button=button)


class MouseClick:
    """鼠标在当前位置点击"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def parsing_ins_dic(self):
        """解析指令字典"""
        return self.ins_dic.get('参数1（键鼠指令）'), \
            int(self.ins_dic.get('参数2').split('-')[0]), \
            int(self.ins_dic.get('参数2').split('-')[1]) / 1000, \
            int(self.ins_dic.get('参数2').split('-')[2]) / 1000

    def start_execute(self):
        """执行重复次数"""
        button_type, click_times, duration, interval = self.parsing_ins_dic()
        re_try = self.ins_dic.get('重复次数')
        # 执行
        if re_try == 1:
            self.simulated_mouse_click(click_times, button_type, duration, interval)
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.simulated_mouse_click(click_times, button_type, duration, interval)
                i += 1
                time.sleep(self.time_sleep)

    def simulated_mouse_click(self, click_times, lOrR, duration, interval):
        """模拟鼠标点击
        :param duration:按压时长,单位：秒
        :param interval:时间间隔,单位：秒
        :param click_times: 点击次数
        :param lOrR: (左键、右键)"""
        button = 'left' if lOrR == '左键' else 'right'
        for i in range(click_times):
            mouse.press(button=button)
            time.sleep(duration)  # 将毫秒转换为秒
            mouse.release(button=button)
            time.sleep(interval)
        self.out_mes.out_mes(f'鼠标在当前位置点击{click_times}次', self.is_test)


class InformationEntry:
    """从Excel中录入信息到窗口"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number
        # 图像点击、文本输入的部分功能
        self.image_click = ImageClick(self.out_mes, self.ins_dic)
        self.text_input = TextInput(self.out_mes, self.ins_dic)

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '点击次数': 3,
            '按钮类型': 'left',
            '工作簿路径': self.ins_dic.get('参数1（键鼠指令）').split('-')[0],
            '工作表名称': self.ins_dic.get('参数1（键鼠指令）').split('-')[1],
            '图像路径': self.ins_dic.get('图像路径'),
            '单元格位置': self.ins_dic.get('参数2'),
            '行号递增': self.ins_dic.get('参数3').split('-')[0],
            '特殊控件输入': self.ins_dic.get('参数3').split('-')[1],
            '超时报错': self.ins_dic.get('参数4'),
            '异常处理': self.ins_dic.get('异常处理')
        }
        return list_dic

    def start_execute(self):
        """执行重复次数"""
        re_try = self.ins_dic.get('重复次数')
        # 执行滚轮滑动
        if re_try == 1:
            self.information_entry()
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.information_entry()
                i += 1
                time.sleep(self.time_sleep)

    def information_entry(self):
        """信息录入"""
        list_dic = self.parsing_ins_dic()
        # 获取excel表格中的值
        cell_value = self.extra_excel_cell_value(
            list_dic.get('工作簿路径'),
            list_dic.get('工作表名称'),
            list_dic.get('单元格位置'),
            eval(list_dic.get('行号递增')),
            self.cycle_number
        )
        self.image_click.execute_click(
            click_times=list_dic.get('点击次数'),
            gray_rec=False,
            lOrR=list_dic.get('按钮类型'),
            img=list_dic.get('图像路径'),
            skip=list_dic.get('超时报错')
        )
        self.text_input.text_input(cell_value, list_dic.get('特殊控件输入'))
        self.out_mes.out_mes('已执行信息录入', self.is_test)

    def extra_excel_cell_value(self,
                               excel_path,
                               sheet_name,
                               cell_position,
                               line_number_increment_,
                               number):
        """获取excel表格中的值
        :param excel_path: excel表格路径
        :param sheet_name: 表格名称
        :param cell_position: 单元格位置
        :param line_number_increment_: 行号递增
        :param number: 循环次数"""
        cell_value = None
        try:
            # 打开excel表格
            wb = openpyxl.load_workbook(excel_path)
            # 选择表格
            sheet = wb[str(sheet_name)]
            if not line_number_increment_:
                # 获取单元格的值
                cell_value = sheet[cell_position].value
                self.out_mes.out_mes(f'获取到的单元格值为：{str(cell_value)}', self.is_test)
            elif line_number_increment_:
                # 获取行号递增的单元格的值
                column_number = re.findall(r"[a-zA-Z]+", cell_position)[0]
                line_number = int(re.findall(r"\d+\.?\d*", cell_position)[0]) + number - 1
                new_cell_position = column_number + str(line_number)
                cell_value = sheet[new_cell_position].value
                self.out_mes.out_mes(f'获取到的单元格值为：{str(cell_value)}', self.is_test)
            return cell_value
        except FileNotFoundError:
            print('没有找到工作簿')
            self.out_mes.out_mes('没有找到工作簿', self.is_test)
            exit_main_work()
        except KeyError:
            print('没有找到工作表')
            self.out_mes.out_mes('没有找到工作表', self.is_test)
            exit_main_work()
        except AttributeError:
            print('没有找到单元格')
            exit_main_work()
            self.out_mes.out_mes('没有找到单元格', self.is_test)


class OpenWeb:
    """打开网页"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        self.out_mes = outputmessage  # 主窗口
        self.ins_dic = ins_dic  # 指令字典
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def start_execute(self):
        """执行重复次数"""
        url = self.ins_dic.get('图像路径')
        global DRIVER
        DRIVER = self.web_option.open_driver(url, True)
        self.out_mes.out_mes('已打开网页', self.is_test)


class EleControl:
    """网页控制"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        self.out_mes = outputmessage  # 主窗口
        self.ins_dic = ins_dic  # 指令字典
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '元素类型': self.ins_dic.get('图像路径').split('-')[0],
            '元素值': self.ins_dic.get('图像路径').split('-')[1],
            '操作类型': self.ins_dic.get('参数1（键鼠指令）'),
            '文本内容': sub_variable(self.ins_dic.get('参数2')),
            '超时类型': self.ins_dic.get('参数3')
        }
        return list_dic

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.text = list_ins_.get('文本内容')
        self.web_option.single_shot_operation(action=list_ins_.get('操作类型'),
                                              element_value_=list_ins_.get('元素值'),
                                              element_type_=list_ins_.get('元素类型'),
                                              timeout_type_=list_ins_.get('超时类型'))
        self.out_mes.out_mes('已执行元素控制', self.is_test)


class WebEntry:
    """将Excel中的值录入网页"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.InformationEntry = InformationEntry(self.out_mes, self.ins_dic)
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '工作簿路径': self.ins_dic.get('参数1（键鼠指令）').split('-')[0],
            '工作表名称': self.ins_dic.get('参数1（键鼠指令）').split('-')[1],
            '元素类型': self.ins_dic.get('图像路径').split('-')[0],
            '元素值': self.ins_dic.get('图像路径').split('-')[1],
            '单元格位置': self.ins_dic.get('参数2'),
            '行号递增': self.ins_dic.get('参数3'),
            '超时类型': self.ins_dic.get('参数4')
        }
        return list_dic

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 获取excel表格中的值
        cell_value = self.InformationEntry.extra_excel_cell_value(
            list_ins_.get('工作簿路径'),
            list_ins_.get('工作表名称'),
            list_ins_.get('单元格位置'),
            bool(list_ins_.get('行号递增')),
            self.cycle_number
        )
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.text = cell_value
        self.out_mes.out_mes('已获取到单元格值', self.is_test)
        self.web_option.single_shot_operation(action='输入内容',
                                              element_value_=list_ins_.get('元素值'),
                                              element_type_=list_ins_.get('元素类型'),
                                              timeout_type_=list_ins_.get('超时类型')
                                              )
        self.out_mes.out_mes('已执行信息录入', self.is_test)


class MouseDrag:
    """鼠标拖拽"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def parsing_ins_dic(self):
        """解析指令字典"""
        start_position = tuple(list(map(int, dict(self.ins_dic)['参数1（键鼠指令）'].split(','))))
        end_position = tuple(list(map(int, dict(self.ins_dic)['参数2'].split(','))))
        return {'起始位置': start_position, '结束位置': end_position}

    def mouse_drag(self, start_position, end_position):
        """鼠标拖拽事件"""
        pyautogui.moveTo(start_position[0], start_position[1], duration=0.3)
        pyautogui.dragTo(end_position[0], end_position[1], duration=0.3)
        self.out_mes.out_mes('鼠标拖拽%s到%s' % (str(start_position), str(end_position)), self.is_test)

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        re_try = self.ins_dic.get('重复次数')
        # 执行滚轮滑动
        if re_try == 1:
            self.mouse_drag(list_ins_.get('起始位置'), list_ins_.get('结束位置'))
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.mouse_drag(list_ins_.get('起始位置'), list_ins_.get('结束位置'))
                time.sleep(self.time_sleep)
                i += 1


class SaveForm:
    """保存网页表格"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)

    def parsing_ins_dic(self):
        """解析指令字典"""
        image_path_parts = dict(self.ins_dic)['图像路径'].split('-')
        element_type, element_value = image_path_parts[0], image_path_parts[1]
        keyboard_mouse_parts = dict(self.ins_dic)['参数1（键鼠指令）'].split('-')
        excel_path, sheet_name = keyboard_mouse_parts[0], keyboard_mouse_parts[1]
        timeout_type = dict(self.ins_dic)['参数2']
        return {
            '元素类型': element_type,
            '元素值': element_value,
            '工作簿路径': excel_path,
            '工作表名称': sheet_name,
            '超时类型': timeout_type
        }

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.single_shot_operation(action='保存表格',
                                              element_value_=list_ins_.get('元素值'),
                                              element_type_=list_ins_.get('元素类型'),
                                              timeout_type_=list_ins_.get('超时类型')
                                              )
        self.out_mes.out_mes('已执行保存网页表格', self.is_test)


class ToggleFrame:
    """切换frame"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '切换类型': self.ins_dic.get('参数1（键鼠指令）'),
            'frame类型': self.ins_dic.get('参数2'),
            'frame值': self.ins_dic.get('参数3'),
        }
        return list_dic

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 执行网页操作
        self.web_option.switch_to_frame(
            iframe_type=list_ins_.get('frame类型'),
            iframe_value=list_ins_.get('frame值'),
            switch_type=list_ins_.get('切换类型')
        )
        self.out_mes.out_mes('已执行切换frame', self.is_test)


class SwitchWindow:
    """切换网页窗口"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '切换类型': self.ins_dic.get('参数1（键鼠指令）'),
            '窗口值': self.ins_dic.get('参数2'),
        }
        return list_dic

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 执行网页操作
        self.web_option.switch_to_window(
            window_type=list_ins_.get('切换类型'),
            window_value=list_ins_.get('窗口值')
        )
        self.out_mes.out_mes('已执行切换窗口', self.is_test)


class DragWebElements:
    """拖拽网页元素"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False
        self.cycle_number = cycle_number
        # 网页控制的部分功能
        self.web_option = WebOption(self.out_mes)

    def parsing_ins_dic(self):
        """解析指令字典"""
        image_path_parts = dict(self.ins_dic)['图像路径'].split('-')
        element_type, element_value = image_path_parts[0], image_path_parts[1]
        x, y = map(int, dict(self.ins_dic)['参数1（键鼠指令）'].split('-'))
        timeout_type = dict(self.ins_dic)['参数2']
        return {
            '元素类型': element_type,
            '元素值': element_value,
            'x': x,
            'y': y,
            '超时类型': timeout_type
        }

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.distance_x = int(dict(list_ins_)['x'])
        self.web_option.distance_y = int(dict(list_ins_)['y'])
        self.web_option.single_shot_operation(action='拖动元素',
                                              element_value_=list_ins_.get('元素值'),
                                              element_type_=list_ins_.get('元素类型'),
                                              timeout_type_=list_ins_.get('超时类型')
                                              )
        self.out_mes.out_mes('已执行拖拽网页元素', self.is_test)


class FullScreenCapture:
    """全屏截图"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 主窗口
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number

    def parsing_ins_dic(self):
        """解析指令字典"""
        return {
            '图像路径': dict(self.ins_dic)['图像路径']
        }

    def start_execute(self):
        """执行重复次数"""
        image_path = self.parsing_ins_dic().get('图像路径', '') + '.png'
        # 执行截图
        screenshot = pyautogui.screenshot()
        # 将图片保存到指定文件夹
        screenshot.save(image_path)
        self.out_mes.out_mes('已执行全屏截图', self.is_test)


# class SendWeChat:
#     """发送微信消息"""
#
#     def __init__(self, outputmessage, ins_dic, cycle_number=1):
#         # 设置参数
#         self.time_sleep = float(get_setting_data_from_db('暂停时间'))
#         self.out_mes = outputmessage
#         # 指令字典
#         self.ins_dic = ins_dic
#         # 是否是测试
#         self.is_test = False
#         self.cycle_number = cycle_number
#
#     def parsing_ins_dic(self):
#         """解析指令字典"""
#         return {
#             '联系人': self.ins_dic.get('参数1（键鼠指令）'),
#             '消息内容': sub_variable(self.ins_dic.get('参数2')),
#         }
#
#     @staticmethod
#     def check_course(title_):
#         """检查软件是否正在运行
#         :param title_: 窗口标题"""
#
#         def get_all_window_title():
#             """获取所有窗口句柄和窗口标题"""
#             hwnd_title_ = dict()
#
#             def get_all_hwnd(hwnd_, mouse):
#                 # print(mouse)
#                 if win32gui.IsWindow(hwnd_) and win32gui.IsWindowEnabled(hwnd_) and win32gui.IsWindowVisible(hwnd_):
#                     hwnd_title_.update({hwnd_: win32gui.GetWindowText(hwnd_)})
#
#             win32gui.EnumWindows(get_all_hwnd, 0)
#             return hwnd_title_
#
#         hwnd_title = get_all_window_title()
#         for h, t in hwnd_title.items():
#             if t == title_:
#                 return h
#
#     def send_message_to_wechat(self, contact_person, message, repeat_times=1):
#         """向微信好友发送消息
#         :param contact_person: 联系人
#         :param message: 消息内容
#         :param repeat_times: 重复次数"""
#
#         def get_process_id(hwnd_):
#             thread_id, process_id_ = win32process.GetWindowThreadProcessId(hwnd_)
#             return process_id_
#
#         def get_correct_message():
#             """获取正确的窗口句柄"""
#             if message == '从剪切板粘贴':
#                 return pyperclip.paste()
#             elif message == '当前日期时间':
#                 return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
#             else:
#                 return message
#
#         def output_info(judge, message_=None, failure_info=None):
#             """向主窗口或na输出提示信息
#             :param failure_info:失败信息
#             :param judge: （成功、失败）
#             :param message_: 消息内容，可选"""
#             output_message = None
#             if judge == '成功':
#                 output_message = f'微信已发送消息：{message_}' if message_ else f'已发送消息'
#             elif judge == '失败':
#                 output_message = f'{failure_info}'
#             self.out_mes.out_mes(output_message, self.is_test)
#
#         pyautogui.hotkey('ctrl', 'alt', 'w')  # 打开微信窗口
#         hwnd = self.check_course('微信')
#         new_message = get_correct_message()
#         try:
#             if hwnd:
#                 process_id = get_process_id(hwnd)  # 获取微信进程id
#                 # 连接到wx
#                 wx_app = Application(backend='uia').connect(process=process_id)
#                 # 定位到主窗口
#                 wx_win = wx_app.window(class_name='WeChatMainWndForPC')
#                 wx_chat_win = wx_win.child_window(title=contact_person, control_type="ListItem")
#                 # 聚焦到所需的对话框
#                 wx_chat_win.click_input()
#
#                 for i in range(repeat_times):  # 重复次数
#                     pyperclip.copy(new_message)  # 将消息内容复制到剪切板
#                     pyautogui.hotkey('ctrl', 'v')
#                     pyautogui.press('enter')  # 模拟按下键盘enter键，发送消息
#                     time.sleep(self.time_sleep)
#
#                 win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  # 最小化窗口
#                 output_info('成功', new_message)  # 向主窗口输出提示信息
#             else:
#                 output_info('失败', new_message, '未找到微信窗口，发送失败。')  # 向主窗口输出提示信息
#         except ElementNotFoundError:
#             win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  # 最小化窗口
#             output_info('失败', new_message, '未找到联系人，发送失败。')  # 向主窗口输出提示信息
#
#     def start_execute(self):
#         """执行重复次数"""
#         list_ins_ = self.parsing_ins_dic()
#         re_try = self.ins_dic.get('重复次数')
#         # 执行滚轮滑动
#         if re_try == 1:
#             self.send_message_to_wechat(list_ins_.get('联系人'), list_ins_.get('消息内容'))
#         elif re_try > 1:
#             self.send_message_to_wechat(list_ins_.get('联系人'), list_ins_.get('消息内容'), re_try)


# class VerificationCode:
#
#     def __init__(self, outputmessage, ins_dic, cycle_number=1):
#         # 主窗口
#         self.out_mes = outputmessage
#         # 指令字典
#         self.ins_dic = ins_dic
#         # 网页控制的部分功能
#         self.web_option = WebOption(self.out_mes)
#         # 是否是测试
#         self.is_test = False
#         self.cycle_number = cycle_number
#
#     def parsing_ins_dic(self):
#         """解析指令字典"""
#         return {
#             '截图区域': self.ins_dic.get('参数1（键鼠指令）'),
#             '元素类型': self.ins_dic.get('参数2'),
#             '元素值': self.ins_dic.get('图像路径'),
#         }
#
#     def ver_input(self, region, element_type, element_value):
#         """截图区域，识别验证码，输入验证码"""
#         im = pyautogui.screenshot(region=(region[0], region[1], region[2], region[3]))
#         im_bytes = io.BytesIO()
#         im.save(im_bytes, format='PNG')
#         im_b = im_bytes.getvalue()
#         ocr = ddddocr.DdddOcr()
#         res = ocr.classification(im_b)
#         self.out_mes.out_mes(f'识别出的验证码为：{res}', self.is_test)
#         # 释放资源
#         del im
#         del im_bytes
#         # 执行网页操作
#         global DRIVER
#         self.web_option.driver = DRIVER
#         self.web_option.text = res
#         self.web_option.single_shot_operation(action='输入内容',
#                                               element_value_=element_value,
#                                               element_type_=element_type,
#                                               timeout_type_=10)
#
#     def start_execute(self):
#         """执行重复次数"""
#         list_dic = self.parsing_ins_dic()
#         verification_code_region = eval(list_dic.get('截图区域'))
#         # 执行验证码输入
#         self.ver_input(
#             verification_code_region,
#             list_dic.get('元素类型'),
#             list_dic.get('元素值')
#         )


class PlayVoice:
    """播放声音"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self, type_):
        """解析指令字典"""
        if type_ == '系统提示音':
            return self.ins_dic.get('参数2')
        elif type_ == '音频信号':
            return int(self.ins_dic.get('参数2').split(',')[0]), \
                int(self.ins_dic.get('参数2').split(',')[1]), \
                int(self.ins_dic.get('参数2').split(',')[2]), \
                int(self.ins_dic.get('参数2').split(',')[3])
        elif type_ == '播放语音':
            return sub_variable(self.ins_dic.get('参数2')), int(self.ins_dic.get('参数3'))

    def play_voice(self, type_):
        """播放声音"""
        if type_ == '系统提示音':
            self.out_mes.out_mes('播放系统提示音', self.is_test)
            self.system_prompt_tone(self.parsing_ins_dic('系统提示音'))
        elif type_ == '音频信号':
            self.out_mes.out_mes('播放音频信号', self.is_test)
            self.sound_signal(*self.parsing_ins_dic('音频信号'))
        elif type_ == '播放语音':
            self.out_mes.out_mes('播放语音', self.is_test)
            self.play_audio(*self.parsing_ins_dic('播放语音'))

    def start_execute(self):
        """开始执行鼠标点击事件"""
        reTry = self.ins_dic.get('重复次数')
        type_ = self.ins_dic.get('参数1（键鼠指令）')
        # 执行坐标点击
        if reTry == 1:
            self.play_voice(type_)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.play_voice(type_)
                i += 1
                time.sleep(self.time_sleep)

    @staticmethod
    def system_prompt_tone(sound_type) -> None:
        """系统提示音
        :param sound_type: 提示音类型(1:警告, 2:错误, 3:询问, 4:信息, 5:系统启动, 6:系统关闭)"""
        if sound_type == '系统警告':
            winsound.PlaySound('SystemAsterisk', winsound.SND_ALIAS)
        elif sound_type == '系统错误':
            winsound.PlaySound('SystemExclamation', winsound.SND_ALIAS)
        elif sound_type == '系统询问':
            winsound.PlaySound('SystemQuestion', winsound.SND_ALIAS)
        elif sound_type == '系统信息':
            winsound.PlaySound('SystemHand', winsound.SND_ALIAS)
        elif sound_type == '系统启动':
            winsound.PlaySound('SystemStart', winsound.SND_ALIAS)
        elif sound_type == '系统关闭':
            winsound.PlaySound('SystemExit', winsound.SND_ALIAS)

    @staticmethod
    def sound_signal(frequency: int,
                     duration: int,
                     times: int = 1,
                     interval: int = 0) -> None:
        """播放音频信号
        :param frequency: 频率(37~32767)
        :param duration: 持续时间(毫秒)
        :param times: 次数
        :param interval: 间隔时间(毫秒)"""
        print('播放音频信号')
        print(frequency, duration, times, interval)
        try:
            for _ in range(times):
                winsound.Beep(frequency, duration)
                if interval:
                    time.sleep(interval / 1000)
        except RuntimeError:
            print('播放音频信号失败')

    @staticmethod
    def play_audio(info: str, rate: int = 200) -> None:
        """播放TTS提示音"""
        try:
            engine = pyttsx4.init()
            engine.setProperty('rate', rate)  # 设置语速
            engine.say(info)
            engine.runAndWait()
        except Exception as e:
            print(e)


class WaitWindow:
    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        self.root = tk.Tk()
        # 设置标签1
        self.font = ("微软雅黑", 12)  # 设置字体为微软雅黑
        self.label = tk.Label(self.root, text="", font=self.font, fg='blue')  # 设置字体和颜色
        self.label.pack(pady=1)
        # 设置标签2
        self.font = ("微软雅黑", 50)  # 设置字体为微软雅黑，字体大小为9
        self.label_2 = tk.Label(self.root, text="", font=self.font, fg="red")  # 设置字体和颜色
        self.label_2.pack(pady=1)
        # 设置按钮
        style = ttk.Style()
        style.configure('RoundedButton.TButton',
                        font=('Arial', 12, 'bold'),
                        borderwidth=0, relief=tk.RAISED)
        self.button = ttk.Button(self.root, text="结束等待",
                                 style='RoundedButton.TButton',
                                 command=self.stop_win)
        self.button.pack(pady=1)
        # 其他设置
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic
        self.cycle_number = cycle_number
        self.count = int(self.ins_dic.get('参数3'))
        self.update_label()
        self.is_test = False
        self.root.protocol("WM_DELETE_WINDOW", self.stop_win)  # 点击关闭按钮时执行
        # 窗口设置
        self.root.geometry("300x200")
        self.root.resizable(False, False)  # 设置禁止调整窗口大小
        self.root.attributes("-toolwindow", 2)
        self.root.attributes("-topmost", True)  # 置顶窗口
        self.root.attributes("-alpha", 0.9)  # 设置透明度
        self.set_screen_to_center()

    def set_screen_to_center(self):
        # 移动窗口到屏幕中央
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width - 300) / 2)
        y = int((screen_height - 200) / 2)
        self.root.geometry("+{}+{}".format(x, y))

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '窗口标题': self.ins_dic.get('参数1（键鼠指令）'),
            '提示信息': self.ins_dic.get('参数2'),
            '等待时间': self.ins_dic.get('参数3')
        }
        return list_dic

    def update_label(self):
        if self.count < 1:
            self.root.destroy()
            self.out_mes.out_mes('已结束等待窗口', self.is_test)
            return

        self.label_2.config(text="{}".format(self.count))
        self.count -= 1
        self.root.after(1000, self.update_label)

    def stop_win(self):
        self.root.destroy()
        self.out_mes.out_mes('已结束等待窗口', self.is_test)

    def start_execute(self):
        self.out_mes.out_mes('正在运行等待窗口...', self.is_test)
        self.root.title(self.ins_dic.get('参数1（键鼠指令）'))  # 窗口标题
        self.count = int(self.ins_dic.get('参数3'))  # 等待时间
        self.label.config(text=self.ins_dic.get('参数2'))
        self.root.mainloop()


class DialogWindow:
    """提示框功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '标题': self.ins_dic.get('参数1（键鼠指令）'),
            '内容': self.ins_dic.get('参数2'),
            '图标': self.ins_dic.get('参数3')
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 解析指令字典
        ins_dic = self.parsing_ins_dic()
        self.alert_dialog_box(ins_dic.get('内容'), ins_dic.get('标题'), ins_dic.get('图标'))

    def alert_dialog_box(self, text, title, icon_):
        """弹出对话框
        :param text: 弹窗内容
        :param title: 弹窗标题
        :param icon_: 弹窗图标"""
        self.out_mes.out_mes('已执行弹窗', self.is_test)
        icon_dic = {
            'STOP': pymsgbox.STOP,
            'WARNING': pymsgbox.WARNING,
            'INFO': pymsgbox.INFO,
            'QUESTION': pymsgbox.QUESTION,
        }
        pymsgbox.alert(
            text=text,
            title=title,
            icon=icon_dic.get(icon_)
        )


class BranchJump:
    """跳转分支的功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典

        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return self.ins_dic.get('重复次数')

    def start_execute(self):
        """开始执行鼠标点击事件"""
        raise IndexError


class TerminationProcess:
    """终止流程的功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典

        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return self.ins_dic.get('重复次数')

    def start_execute(self):
        """开始执行鼠标点击事件"""
        raise IndexError


class WindowControl:
    """窗口控制"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典

        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '窗口标题': self.ins_dic.get('参数1（键鼠指令）'),
            '操作类型': self.ins_dic.get('参数2').split('-')[0].replace('窗口', ''),
            '是否报错': eval(self.ins_dic.get('参数2').split('-')[1])
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        list_dic = self.parsing_ins_dic()
        self.show_normal_window_with_specified_title(
            list_dic.get('窗口标题'),
            list_dic.get('操作类型'),
            list_dic.get('是否报错')
        )

    def show_normal_window_with_specified_title(self, title, judge='最大化', is_error=True):
        """将指定标题的窗口置顶
        :param is_error: 是否报错
        :param title: 指定标题
        :param judge: 判断（最大化、最小化、显示窗口、关闭）"""

        def get_all_window_title():
            """获取所有窗口句柄和窗口标题"""
            hwnd_title_ = dict()

            def get_all_hwnd(hwnd, mouse):
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                    hwnd_title_.update({hwnd: win32gui.GetWindowText(hwnd)})

            win32gui.EnumWindows(get_all_hwnd, 0)
            return hwnd_title_

        hwnd_title = get_all_window_title()
        for h, t in hwnd_title.items():
            if title in t:
                if judge == '最大化':
                    win32gui.ShowWindow(h, win32con.SW_SHOWMAXIMIZED)  # 最大化显示窗口
                elif judge == '最小化':
                    win32gui.ShowWindow(h, win32con.SW_SHOWMINIMIZED)
                elif judge == '显示':
                    win32gui.ShowWindow(h, win32con.SW_SHOWNORMAL)  # 显示窗口
                elif judge == '关闭':
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                self.out_mes.out_mes(f'已{judge}指定标题包含“{title}”的窗口', self.is_test)
                break
        else:
            self.out_mes.out_mes(f'没有找到标题包含“{title}”的窗口！', self.is_test)
            if is_error:
                raise ValueError(f'没有找到标题包含“{title}”的窗口！')


class KeyWait:
    """按键等待的功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典

        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

        self.key: str = ''

    def start_execute(self):
        """开始执行鼠标点击事件"""
        self.key = self.ins_dic.get('参数1（键鼠指令）')
        type_ = self.ins_dic.get('参数2')
        self.out_mes.out_mes(f'等待按键{self.key}按下中...', self.is_test)
        if type_ == '等待按键':
            keyboard.wait(self.key.lower())
            self.out_mes.out_mes(f'按键{self.key}已被按下', self.is_test)
        elif type_ == '等待跳转分支':
            keyboard.wait(self.key.lower())
            self.out_mes.out_mes(f'按键{self.key}已被按下！跳转分支。', self.is_test)
            raise ValueError(f'按键{self.key}已被按下！')


class GetTimeValue:
    """获取时间变量指令"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '时间格式': self.ins_dic.get('参数1（键鼠指令）'),
            '变量名称': self.ins_dic.get('参数2')
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        list_dic = self.parsing_ins_dic()  # 参数字典
        now_time_str = str(self.get_now_time(list_dic.get('时间格式')))
        if not self.is_test:
            set_variable_value(list_dic.get('变量名称'), now_time_str)
            self.out_mes.out_mes(f'已获取当前时间并赋值给变量：{list_dic.get("变量名称")}', self.is_test)
        else:
            self.out_mes.out_mes(f'已获取当前时间：{now_time_str}', self.is_test)

    @staticmethod
    def get_now_time(format_="年-月-日 小时:分钟:秒"):
        """获取当前时间
        :param format_: 时间格式
        :return: 当前时间字符串"""
        allowed_formats = {
            "年-月-日 小时:分钟:秒": "%Y-%m-%d %H:%M:%S",
            "年/月/日 小时:分钟:秒": "%Y/%m/%d %H:%M:%S",
            "月/日/年 小时:分钟:秒": "%m/%d/%Y %H:%M:%S",
            "日-月-年 小时:分钟:秒": "%d-%m-%Y %H:%M:%S",
            "年-月-日": "%Y-%m-%d",
            "月/日/年": "%m/%d/%Y",
            "日-月-年": "%d-%m-%Y",
            "年-月": "%Y-%m",
            "月/年": "%m/%Y",
            "年": "%Y",
            "时间戳": "%s"  # 时间戳格式
        }

        if format_ not in allowed_formats:
            raise ValueError("Invalid format_. "
                             "Please use one of the allowed formats.")
        if format_ == "时间戳":
            return int(time.time())
        else:
            return time.strftime(allowed_formats[format_], time.localtime())


class GetExcelCellValue:
    """从excel表格中获取单元格的值"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '工作簿路径': self.ins_dic.get('图像路径').split('-')[0],
            '工作表名称': self.ins_dic.get('图像路径').split('-')[1],
            '单元格位置': self.ins_dic.get('参数1（键鼠指令）'),
            '变量名称': self.ins_dic.get('参数2'),
            '行号递增': eval(self.ins_dic.get('参数3'))
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        list_dic = self.parsing_ins_dic()
        cell_position = list_dic.get('单元格位置')

        if list_dic.get('行号递增'):  # 递增行号
            cell_position = line_number_increment(cell_position, self.cycle_number - 1)

        cell_value = self.get_value_from_excel(
            list_dic.get('工作簿路径'),
            list_dic.get('工作表名称'),
            cell_position
        )
        self.send_out_message(cell_value, list_dic)

    def send_out_message(self, cell_value, list_dic):
        if not self.is_test:
            set_variable_value(list_dic.get('变量名称'), cell_value)
            self.out_mes.out_mes(f'已获取单元格的值并赋值给变量：{list_dic.get("变量名称")}', self.is_test)
        else:
            self.out_mes.out_mes(f'已获取单元格的值：{cell_value}', self.is_test)

    @staticmethod
    def get_value_from_excel(file_path, sheet_name, cell='A1'):
        """从excel中获取数据"""
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb[sheet_name]
            value = sheet[cell].value
            return value
        except Exception as e:
            print(e)
            return None


class GetDialogValue:
    """从对话框中获取值"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '对话框标题': self.ins_dic.get('参数1（键鼠指令）'),
            '变量名称': self.ins_dic.get('参数2'),
            '对话框提示信息': self.ins_dic.get('参数3')
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        ins_dic = self.parsing_ins_dic()  # 解析指令字典
        text = self.gets_text_from_dialog(ins_dic)
        set_variable_value(ins_dic.get('变量名称'), text)
        self.out_mes.out_mes(f'已获取对话框的值并赋值给变量：{ins_dic.get("变量名称")}', self.is_test)

    @staticmethod
    def gets_text_from_dialog(ins_dic):
        return pymsgbox.prompt(
            ins_dic.get('对话框提示信息'),
            ins_dic.get('对话框标题')
        )


class ContrastVariables:
    """变量判断的功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '变量1': self.ins_dic.get('参数1（键鼠指令）').split('-')[0],
            '比较符': self.ins_dic.get('参数2').split('-')[0],
            '变量2': self.ins_dic.get('参数1（键鼠指令）').split('-')[1],
            '变量类型': self.ins_dic.get('参数2').split('-')[1]
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        ins_dic = self.parsing_ins_dic()
        variable_dic = get_variable_info('dict')  # 获取变量字典
        # 获取变量名称
        variable1_name = ins_dic.get('变量1')
        variable2_name = ins_dic.get('变量2')
        variable_symbol = ins_dic.get('比较符')
        # 获取变量值
        variable1 = variable_dic.get(variable1_name)
        variable2 = variable_dic.get(variable2_name)
        # 执行变量判断
        result = self.comparison_variable(
            variable1, variable_symbol, variable2, ins_dic.get('变量类型')
        )
        # 输出信息
        self.out_mes.out_mes(
            f'变量判断"{variable1_name}{variable_symbol}{variable2_name}"结果：{result}',
            self.is_test
        )
        if result:
            raise ValueError('变量判断结果为真，跳转分支。')

    @staticmethod
    def comparison_variable(variable1,
                            comparison_symbol,
                            variable2,
                            variable_type):
        """比较变量"""

        def try_parse_date(variable):
            """尝试将变量解析为日期时间对象"""
            try:
                return parse(variable)
            except ValueError:
                return None

        variable1_ = variable1
        variable2_ = variable2
        if variable_type == '日期或时间':
            variable1_ = try_parse_date(variable1)
            variable2_ = try_parse_date(variable2)
        elif variable_type == '数字':
            variable1_ = eval(variable1)
            variable2_ = eval(variable2)
        elif variable_type == '字符串':
            variable1_ = str(variable1)
            variable2_ = str(variable2)

        if comparison_symbol == '=':
            return variable1_ == variable2_
        elif comparison_symbol == '≠':
            return variable1_ != variable2_
        elif comparison_symbol == '>':
            return variable1_ > variable2_
        elif comparison_symbol == '<':
            return variable1_ < variable2_
        elif comparison_symbol == '包含':
            return variable1_ in variable2_


class RunPython:
    """运行python脚本"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '返回名称': self.ins_dic.get('参数1（键鼠指令）'),
            '变量名称': self.ins_dic.get('参数2'),
            '代码': self.ins_dic.get('参数3')
        }

    @staticmethod
    def sub_variable_2(text: str):
        """将text中的变量替换为变量值"""
        new_text = text
        if ('☾' in text) and ('☽' in text):
            variable_dic = get_variable_info('dict')
            for key, value in variable_dic.items():
                new_text = new_text.replace(f'☾{key}☽', str(f'"{value}"'))
        return new_text

    def start_execute(self):
        """开始执行鼠标点击事件"""
        ins_dic = self.parsing_ins_dic()
        try:
            # 定义全局命名空间字典
            globals_dict = {}
            python_code = self.sub_variable_2(ins_dic.get('代码'))
            # 在执行代码时，将结果保存到全局命名空间中
            exec(python_code, globals_dict)
            # 从全局命名空间中获取结果
            result = globals_dict.get(ins_dic.get('返回名称'), None)
            if result is not None:
                if not self.is_test:  # 不是测试时,将结果赋值给变量
                    set_variable_value(ins_dic.get('变量名称'), result)
            self.out_mes.out_mes(f'已执行Python，返回：{result}', self.is_test)
        except Exception as e:
            print(e)
            self.out_mes.out_mes(f'运行失败：{e}', self.is_test)


class RunExternalFile:
    """运行外部文件"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return self.ins_dic.get('图像路径')

    def start_execute(self):
        """开始执行鼠标点击事件"""
        file_path = self.parsing_ins_dic()
        self.run_external_file(file_path)

    def run_external_file(self, file_path):
        """运行外部文件"""
        try:
            os.startfile(file_path)
            self.out_mes.out_mes(f'已运行外部文件：{file_path}', self.is_test)
        except Exception as e:
            print(e)
            self.out_mes.out_mes(f'运行失败：{e}', self.is_test)
            raise ValueError(f'打开文件失败')


class InputCellExcel:
    """输入到excel单元格的功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '工作簿路径': self.ins_dic.get('图像路径').split('-')[0],
            '工作表名称': self.ins_dic.get('图像路径').split('-')[1],
            '单元格位置': self.ins_dic.get('参数1（键鼠指令）').split('-')[0],
            '是否递增': eval(self.ins_dic.get('参数1（键鼠指令）').split('-')[1]),
            '输入内容': self.ins_dic.get('参数2')
        }

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 解析指令字典
        list_dic = self.parsing_ins_dic()
        excel_path = list_dic.get('工作簿路径')
        sheet_name = list_dic.get('工作表名称')
        cell_position = list_dic.get('单元格位置')

        if list_dic.get('是否递增'):
            cell_position = line_number_increment(cell_position, self.cycle_number - 1)

        # 输入到excel单元格
        self.input_to_excel(
            excel_path,
            sheet_name,
            cell_position,
            sub_variable(list_dic.get('输入内容'))
        )

    def input_to_excel(self,
                       file_path,
                       sheet_name,
                       cell_position,
                       input_content
                       ):
        """输入到excel单元格
        :param file_path: 文件路径
        :param sheet_name: 表名
        :param cell_position: 单元格位置
        :param input_content: 输入内容"""
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb[sheet_name]
            sheet[cell_position] = input_content
            wb.save(file_path)
            self.out_mes.out_mes(f'已将"{input_content}"输入到单元格"{cell_position}"', self.is_test)
        except Exception as e:
            print(e)
            self.out_mes.out_mes(f'输入失败：{e}', self.is_test)
            raise ValueError(f'输入失败')


class TextRecognition:
    """文字识别功能"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数
        self.transparent_window = TransparentWindow()  # 框选窗口

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        return {
            '重复次数': self.ins_dic.get('重复次数'),
            '截图区域': self.ins_dic.get('参数1（键鼠指令）'),
            '变量名称': self.ins_dic.get('参数2')
        }

    def start_execute(self):
        """开始执行事件"""
        list_dic = self.parsing_ins_dic()
        ocr_text = self.ocr_pic(list_dic['截图区域'])  # 识别图片中的文字
        # 显示识别结果
        if (ocr_text is not None) and (ocr_text != ''):
            self.out_mes.out_mes(f'OCR识别结果：{ocr_text}', self.is_test)
            if not self.is_test:  # 如果不是测试
                set_variable_value(list_dic['变量名称'], ocr_text)
                self.out_mes.out_mes(
                    f'已将OCR识别结果赋值给变量：{list_dic["变量名称"]}', self.is_test
                )
        else:
            self.out_mes.out_mes('OCR识别失败！检查网络或查看OCR信息是否设置正确。', self.is_test)

    @staticmethod
    def ocr_pic(reigon):
        """文字识别
        :param reigon: 识别区域"""

        def get_result_from_text(text):
            """从识别结果中提取文字信息"""
            return '\n'.join(i['words'] for i in text.get('words_result', []))

        im = pyautogui.screenshot(region=eval(reigon))
        # 将截图数据存储在内存中
        im_bytes = io.BytesIO()
        im.save(im_bytes, format='PNG')
        im_b = im_bytes.getvalue()
        # 返回百度api识别文字信息
        try:
            client_info = get_ocr_info()  # 获取百度api信息
            client = AipOcr(client_info['appId'], client_info['apiKey'], client_info['secretKey'])
            return get_result_from_text(client.basicGeneral(im_b))
        except Exception as e:
            print(f'Error: {e} 网络错误识别失败')
            return None
        finally:  # 释放内存
            del im
            del im_bytes
