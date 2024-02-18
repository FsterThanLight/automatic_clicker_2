import io
import random
import re
import sys
import time
from datetime import datetime

import ddddocr
import keyboard
import mouse
import openpyxl
import pyautogui
import pyperclip
import win32con
import win32gui
import win32process
from PyQt5.QtWidgets import QApplication
from dateutil.parser import parse

from 数据库操作 import get_setting_data_from_db, get_str_now_time
from 网页操作 import WebOption

sys.coinit_flags = 2  # STA
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

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


class OutputMessage:
    """输出信息，测试时输出到文本框，非测试时输出到主窗口"""

    def __init__(self, command_thread, navigation):
        # 输出的窗口
        self.command_thread = command_thread
        self.navigation = navigation

    def out_mes(self, message, is_test: bool = False):
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

    def __init__(self, outputmessage, ins_dic):
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

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数
        :return: 指令参数列表，重复次数"""
        # 读取图像名称
        img = self.ins_dic.get('图像路径')
        # 取重复次数
        re_try = self.ins_dic.get('重复次数')
        skip = self.ins_dic.get('参数2')  # 是否跳过参数
        gray_recognition = eval(self.ins_dic.get('参数3'))  # 是否灰度识别
        click_map = {
            '左键单击': [1, 'left', img, skip],
            '左键双击': [2, 'left', img, skip],
            '右键单击': [1, 'right', img, skip],
            '右键双击': [2, 'right', img, skip],
            '仅移动鼠标': [0, 'left', img, skip]
        }
        list_ins = click_map.get(self.ins_dic.get('参数1（键鼠指令）'))
        # 返回重复次数，点击次数，左键右键，图片名称，是否跳过
        return re_try, gray_recognition, list_ins[0], list_ins[1], list_ins[2], list_ins[3]

    def start_execute(self, number):
        """开始执行鼠标点击事件
        :param number: 主窗口显示的循环次数"""
        # 解析指令字典
        reTry, gray_rec, click_times, lOrR, img, skip = self.parsing_ins_dic()
        # 执行图像点击
        if reTry == 1:
            self.execute_click(click_times, gray_rec, lOrR, img, skip, number)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.execute_click(click_times, gray_rec, lOrR, img, skip, number)
                i += 1
                time.sleep(self.time_sleep)

    def execute_click(self, click_times, gray_rec, lOrR, img, skip, number):
        """执行鼠标点击事件"""

        # 4个参数：鼠标点击时间，按钮类型（左键右键中键），图片名称，重复次数
        def image_match_click(location):
            if location is not None:
                if not self.is_test:
                    # 参数：位置X，位置Y，点击次数，时间间隔，持续时间，按键
                    self.out_mes.out_mes(f'已找到匹配图片%s' % str(number), self.is_test)
                    pyautogui.click(location.x, location.y,
                                    clicks=click_times,
                                    interval=self.interval,
                                    duration=self.duration,
                                    button=lOrR)
                elif self.is_test:
                    self.out_mes.out_mes(f'已找到匹配图片%s' % str(number), self.is_test)
                    # 移动鼠标到图片位置
                    pyautogui.moveTo(location.x, location.y, duration=0.2)

        try:
            min_search_time = 1 if skip == "自动略过" else float(skip)
            # 显示信息
            self.out_mes.out_mes(f'正在查找匹配图像...', self.is_test)
            QApplication.processEvents()
            location_ = pyautogui.locateCenterOnScreen(
                image=img,
                confidence=self.confidence,
                minSearchTime=min_search_time,
                grayscale=gray_rec
            )
            if location_:  # 如果找到图像
                image_match_click(location_)
            elif not location_:  # 如果未找到图像
                self.out_mes.out_mes('未找到匹配图像', self.is_test)
            QApplication.processEvents()
        except OSError:
            self.out_mes.out_mes('文件下未找到png图像，请检查文件是否存在！', self.is_test)
            raise FileNotFoundError


class CoordinateClick:
    """坐标点击"""

    def __init__(self, outputmessage, ins_dic):
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

    def __init__(self, command_thread, ins_dic):
        # 执行线程
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

    def start_execute(self):
        """从指令字典中解析出指令参数"""
        wait_type = self.ins_dic.get('参数1（键鼠指令）')
        if wait_type == '时间等待':
            wait_time = int(self.ins_dic.get('参数2'))
            self.command_thread.show_message('等待时长%d秒' % wait_time)
            self.stop_time(wait_time)
        elif wait_type == '定时等待':
            target_time, interval_time = self.ins_dic.get('参数2').split('+')
            # 检查目标时间是否大于当前时间
            if parse(target_time) > datetime.now():
                self.wait_to_time(target_time, interval_time)
        elif wait_type == '随机等待':
            min_time, max_time = self.ins_dic.get('参数2').split('-')
            wait_time = random.randint(int(min_time), int(max_time))
            self.command_thread.show_message('随机等待时长%d秒' % wait_time)
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
                self.command_thread.show_message('当前为：%s' % now.strftime('%Y/%m/%d %H:%M:%S'))
                self.command_thread.show_message('等待至：%s' % target_time)
                show_times = sleep_time
            if now >= parse(target_time):
                self.command_thread.show_message('退出等待')
                break
            # 时间暂停
            time.sleep(sleep_time)
            show_times += sleep_time

    def stop_time(self, seconds):
        """暂停时间"""
        for i in range(seconds):
            # 显示剩下等待时间
            self.command_thread.show_message('等待中...剩余%d秒' % (seconds - i))
            time.sleep(1)


class ImageWaiting:
    """图片等待"""

    def __init__(self, command_thread, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

    def wait_to_image(self, image, wait_instruction_type, timeout_period):
        """执行图片等待"""
        if wait_instruction_type == '等待到指定图像出现':
            self.command_thread.show_message('正在等待指定图像出现中...')
            QApplication.processEvents()
            location = pyautogui.locateCenterOnScreen(
                image=image,
                confidence=0.8,
                minSearchTime=timeout_period
            )
            if location:
                self.command_thread.show_message('目标图像已经出现，等待结束')
                QApplication.processEvents()
        elif wait_instruction_type == '等待到指定图像消失':
            vanish = True
            while vanish:
                try:
                    location = pyautogui.locateCenterOnScreen(
                        image=image,
                        confidence=0.8,
                        minSearchTime=1
                    )
                    print('location', location)
                except pyautogui.ImageNotFoundException:
                    self.command_thread.show_message('目标图像已经消失，等待结束')
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

    def __init__(self, outputmessage, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试

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

    def __init__(self, command_thread, ins_dic):
        # 设置参数
        setting_data_dic = get_setting_data_from_db('时间间隔', '暂停时间')
        self.interval = float(setting_data_dic.get('时间间隔'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

    def start_execute(self):
        """解析指令字典"""
        input_value = self.ins_dic.get('参数1（键鼠指令）')
        special_control_judgment = bool(self.ins_dic.get('参数2'))
        # 执行文本输入
        self.text_input(input_value, special_control_judgment)

    def text_input(self, input_value, special_control_judgment):
        """文本输入事件
        :param input_value: 输入的文本
        :param special_control_judgment: 是否为特殊控件"""
        if special_control_judgment == 'False':
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(self.time_sleep)
            self.command_thread.show_message('执行文本输入%s' % input_value)
        elif special_control_judgment == 'True':
            pyautogui.typewrite(input_value, interval=self.interval)
            self.command_thread.show_message('执行特殊控件的文本输入%s' % input_value)
            time.sleep(self.time_sleep)


class MoveMouse:
    """移动鼠标"""

    def __init__(self, outputmessage, ins_dic):
        # 设置参数
        setting_data_dic = get_setting_data_from_db('持续时间', '暂停时间')
        self.duration = float(setting_data_dic.get('持续时间'))
        self.time_sleep = float(setting_data_dic.get('暂停时间'))
        self.out_mes = outputmessage  # 用于输出信息
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试

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

    def __init__(self, command_thread, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

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
        self.command_thread.show_message('按下按键%s' % key)


class MiddleActivation:
    """鼠标中键激活"""

    def __init__(self, command_thread, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

    def start_execute(self):
        """执行重复次数"""
        command_type = self.ins_dic.get('参数1（键鼠指令）')
        click_count = self.ins_dic.get('参数2')
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
        self.command_thread.show_message('等待按下鼠标中键中...按下F11键退出')
        QApplication.processEvents()
        # print('等待按下鼠标中键中...按下esc键退出')
        mouse.wait(button='middle')
        try:
            if command_type == "模拟点击":
                pyautogui.click(clicks=int(click_times), button='left')
                self.command_thread.show_message('执行鼠标点击%d次' % click_times)
                # print('执行鼠标点击' + click_times + '次')
            elif command_type == "自定义":
                pass
        except OSError:
            # 弹出提示框。提示检查鼠标是否连接
            self.command_thread.show_message('连接失败，请检查鼠标是否连接正确。')
            # print('连接失败，请检查鼠标是否连接正确。')


class MouseClick:
    """鼠标在当前位置点击"""

    def __init__(self, outputmessage, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.out_mes = outputmessage
        # 指令字典
        self.ins_dic = ins_dic
        self.is_test = False  # 是否测试

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

    def __init__(self, command_thread, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic
        # 图像点击、文本输入的部分功能
        self.image_click = ImageClick(self.command_thread, self.ins_dic)
        self.text_input = TextInput(self.command_thread, self.ins_dic)

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

    def start_execute(self, number):
        """执行重复次数"""
        re_try = self.ins_dic.get('重复次数')
        # 执行滚轮滑动
        if re_try == 1:
            self.information_entry(number)
        elif re_try > 1:
            i = 1
            while i < re_try + 1:
                self.information_entry(number)
                i += 1
                time.sleep(self.time_sleep)

    def information_entry(self, number):
        """信息录入"""
        list_dic = self.parsing_ins_dic()
        # 获取excel表格中的值
        cell_value = self.extra_excel_cell_value(
            list_dic.get('工作簿路径'),
            list_dic.get('工作表名称'),
            list_dic.get('单元格位置'),
            bool(list_dic.get('行号递增')),
            number
        )
        self.image_click.execute_click(
            click_times=list_dic.get('点击次数'),
            gray_rec=False,
            lOrR=list_dic.get('按钮类型'),
            img=list_dic.get('图像路径'),
            skip=list_dic.get('超时报错'),
            number=number
        )
        self.text_input.text_input(cell_value, list_dic.get('特殊控件输入'))
        self.command_thread.show_message('已执行信息录入')

    def extra_excel_cell_value(self, excel_path, sheet_name,
                               cell_position, line_number_increment, number):
        """获取excel表格中的值
        :param excel_path: excel表格路径
        :param sheet_name: 表格名称
        :param cell_position: 单元格位置
        :param line_number_increment: 行号递增
        :param number: 循环次数"""
        print('正在获取单元格值')
        cell_value = None
        try:
            # 打开excel表格
            wb = openpyxl.load_workbook(excel_path)
            # 选择表格
            sheet = wb[str(sheet_name)]
            if not line_number_increment:
                # 获取单元格的值
                cell_value = sheet[cell_position].value
                self.command_thread.show_message(f'获取到的单元格值为：{str(cell_value)}')
            elif line_number_increment:
                # 获取行号递增的单元格的值
                column_number = re.findall(r"[a-zA-Z]+", cell_position)[0]
                line_number = int(re.findall(r"\d+\.?\d*", cell_position)[0]) + number - 1
                new_cell_position = column_number + str(line_number)
                cell_value = sheet[new_cell_position].value
                self.command_thread.show_message(f'获取到的单元格值为：{str(cell_value)}')
            return cell_value
        except FileNotFoundError:
            print('没有找到工作簿')
            self.command_thread.show_message('没有找到工作簿')
            exit_main_work()
        except KeyError:
            print('没有找到工作表')
            self.command_thread.show_message('没有找到工作表')
            exit_main_work()
        except AttributeError:
            print('没有找到单元格')
            exit_main_work()
            self.command_thread.show_message('没有找到单元格')


class OpenWeb:
    """打开网页"""

    def __init__(self, command_thread, navigation, ins_dic):
        self.command_thread = command_thread  # 主窗口
        self.navigation = navigation  # 导航
        self.ins_dic = ins_dic  # 指令字典
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

    def start_execute(self):
        """执行重复次数"""
        url = self.ins_dic.get('图像路径')
        global DRIVER
        DRIVER = self.web_option.open_driver(url, True)


class EleControl:
    """网页控制"""

    def __init__(self, command_thread, navigation, ins_dic):
        self.command_thread = command_thread  # 主窗口
        self.navigation = navigation  # 导航
        self.ins_dic = ins_dic  # 指令字典
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '元素类型': self.ins_dic.get('图像路径').split('-')[0],
            '元素值': self.ins_dic.get('图像路径').split('-')[1],
            '操作类型': self.ins_dic.get('参数1（键鼠指令）'),
            '文本内容': self.ins_dic.get('参数2'),
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


class WebEntry:
    """将Excel中的值录入网页"""

    def __init__(self, command_thread, navigation, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        self.InformationEntry = InformationEntry(self.command_thread, self.ins_dic)
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

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

    def start_execute(self, number):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        # 获取excel表格中的值
        cell_value = self.InformationEntry.extra_excel_cell_value(
            list_ins_.get('工作簿路径'),
            list_ins_.get('工作表名称'),
            list_ins_.get('单元格位置'),
            bool(list_ins_.get('行号递增')),
            number
        )
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.text = cell_value
        self.web_option.single_shot_operation(action='输入内容',
                                              element_value_=list_ins_.get('元素值'),
                                              element_type_=list_ins_.get('元素类型'),
                                              timeout_type_=list_ins_.get('超时类型')
                                              )


class MouseDrag:
    """鼠标拖拽"""

    def __init__(self, command_thread, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

    def parsing_ins_dic(self):
        """解析指令字典"""
        start_position = tuple(list(map(int, dict(self.ins_dic)['参数1（键鼠指令）'].split(','))))
        end_position = tuple(list(map(int, dict(self.ins_dic)['参数2'].split(','))))
        return {'起始位置': start_position, '结束位置': end_position}

    def mouse_drag(self, start_position, end_position):
        """鼠标拖拽事件"""
        pyautogui.moveTo(start_position[0], start_position[1], duration=0.3)
        pyautogui.dragTo(end_position[0], end_position[1], duration=0.3)
        self.command_thread.show_message('鼠标拖拽%s到%s' % (str(start_position), str(end_position)))

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

    def __init__(self, command_thread, navigation, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

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


class ToggleFrame:
    """切换frame"""

    def __init__(self, command_thread, navigation, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

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


class SwitchWindow:
    """切换网页窗口"""

    def __init__(self, command_thread, navigation, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

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


class DragWebElements:
    """拖拽网页元素"""

    def __init__(self, command_thread, navigation, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 网页控制的部分功能
        self.web_option = WebOption(self.command_thread, self.navigation)

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


class FullScreenCapture:
    """全屏截图"""

    def __init__(self, command_thread, ins_dic):
        # 主窗口
        self.command_thread = command_thread
        # 指令字典
        self.ins_dic = ins_dic

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
        self.command_thread.show_message('已执行全屏截图')


class SendWeChat:
    """发送微信消息"""

    def __init__(self, command_thread, navigation, ins_dic):
        # 设置参数
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        # 主窗口
        self.command_thread = command_thread
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 是否是测试
        self.is_test = False

    def parsing_ins_dic(self):
        """解析指令字典"""
        return {
            '联系人': self.ins_dic.get('参数1（键鼠指令）'),
            '消息内容': self.ins_dic.get('参数2'),
        }

    @staticmethod
    def check_course(title_):
        """检查软件是否正在运行
        :param title_: 窗口标题"""

        def get_all_window_title():
            """获取所有窗口句柄和窗口标题"""
            hwnd_title_ = dict()

            def get_all_hwnd(hwnd_, mouse):
                # print(mouse)
                if win32gui.IsWindow(hwnd_) and win32gui.IsWindowEnabled(hwnd_) and win32gui.IsWindowVisible(hwnd_):
                    hwnd_title_.update({hwnd_: win32gui.GetWindowText(hwnd_)})

            win32gui.EnumWindows(get_all_hwnd, 0)
            return hwnd_title_

        hwnd_title = get_all_window_title()
        for h, t in hwnd_title.items():
            if t == title_:
                return h

    def send_message_to_wechat(self, contact_person, message, repeat_times=1):
        """向微信好友发送消息
        :param contact_person: 联系人
        :param message: 消息内容
        :param repeat_times: 重复次数"""

        def get_process_id(hwnd_):
            thread_id, process_id_ = win32process.GetWindowThreadProcessId(hwnd_)
            return process_id_

        def get_correct_message():
            """获取正确的窗口句柄"""
            if message == '从剪切板粘贴':
                return pyperclip.paste()
            elif message == '当前日期时间':
                return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            else:
                return message

        def output_info(judge, message_=None, failure_info=None):
            """向主窗口或na输出提示信息
            :param failure_info:失败信息
            :param judge: （成功、失败）
            :param message_: 消息内容，可选"""
            output_message = None
            time_now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            if judge == '成功':
                output_message = f'{time_now} 微信已发送消息：{message_}' if message_ else f'{time_now} 已发送消息'
            elif judge == '失败':
                output_message = f'{time_now} {failure_info}'
            if self.is_test:
                self.navigation.textBrowser.append(output_message)
            else:
                self.command_thread.show_message(output_message)

        pyautogui.hotkey('ctrl', 'alt', 'w')  # 打开微信窗口
        hwnd = self.check_course('微信')
        new_message = get_correct_message()
        try:
            if hwnd:
                process_id = get_process_id(hwnd)  # 获取微信进程id
                # 连接到wx
                wx_app = Application(backend='uia').connect(process=process_id)
                # 定位到主窗口
                wx_win = wx_app.window(class_name='WeChatMainWndForPC')
                wx_chat_win = wx_win.child_window(title=contact_person, control_type="ListItem")
                # 聚焦到所需的对话框
                wx_chat_win.click_input()

                for i in range(repeat_times):  # 重复次数
                    pyperclip.copy(new_message)  # 将消息内容复制到剪切板
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press('enter')  # 模拟按下键盘enter键，发送消息
                    time.sleep(self.time_sleep)

                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  # 最小化窗口
                output_info('成功', new_message)  # 向主窗口输出提示信息
            else:
                output_info('失败', new_message, '未找到微信窗口，发送失败。')  # 向主窗口输出提示信息
        except ElementNotFoundError:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  # 最小化窗口
            output_info('失败', new_message, '未找到联系人，发送失败。')  # 向主窗口输出提示信息

    def start_execute(self):
        """执行重复次数"""
        list_ins_ = self.parsing_ins_dic()
        re_try = self.ins_dic.get('重复次数')
        # 执行滚轮滑动
        if re_try == 1:
            self.send_message_to_wechat(list_ins_.get('联系人'), list_ins_.get('消息内容'))
        elif re_try > 1:
            self.send_message_to_wechat(list_ins_.get('联系人'), list_ins_.get('消息内容'), re_try)


class VerificationCode:

    def __init__(self, main_window, navigation, ins_dic):
        # 主窗口
        self.main_window = main_window
        self.navigation = navigation
        # 指令字典
        self.ins_dic = ins_dic
        # 网页控制的部分功能
        self.web_option = WebOption(self.main_window, self.navigation)
        # 是否是测试
        self.is_test = False

    def parsing_ins_dic(self):
        """解析指令字典"""
        return {
            '截图区域': self.ins_dic.get('参数1（键鼠指令）'),
            '元素类型': self.ins_dic.get('参数2'),
            '元素值': self.ins_dic.get('图像路径'),
        }

    def ver_input(self, region, element_type, element_value):
        """截图区域，识别验证码，输入验证码"""
        im = pyautogui.screenshot(region=(region[0], region[1], region[2], region[3]))
        im_bytes = io.BytesIO()
        im.save(im_bytes, format='PNG')
        im_b = im_bytes.getvalue()
        ocr = ddddocr.DdddOcr()
        res = ocr.classification(im_b)
        self.main_window.plainTextEdit.appendPlainText(f'识别出的验证码为：{res}')
        # 释放资源
        del im
        del im_bytes
        # 执行网页操作
        global DRIVER
        self.web_option.driver = DRIVER
        self.web_option.text = res
        self.web_option.single_shot_operation(action='输入内容',
                                              element_value_=element_value,
                                              element_type_=element_type,
                                              timeout_type_=10)

    def start_execute(self):
        """执行重复次数"""
        list_dic = self.parsing_ins_dic()
        verification_code_region = eval(list_dic.get('截图区域'))
        # 执行验证码输入
        self.ver_input(
            verification_code_region,
            list_dic.get('元素类型'),
            list_dic.get('元素值')
        )


if __name__ == '__main__':
    pass
    elem = (1, None, '等待', '等待到指定时间', '2023/10/5 22:46:24+1000', None, None, 1, '抛出异常并暂停', '', '主流程')
    dic = {
        'ID': elem[0],
        '图像路径': elem[1],
        '指令类型': elem[2],
        '参数1（键鼠指令）': elem[3],
        '参数2': elem[4],
        '参数3': elem[5],
        '参数4': elem[6],
        '重复次数': elem[7],
        '异常处理': elem[8]
    }
    main_window_ = None
