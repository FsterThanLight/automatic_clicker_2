# Copyright (c) [2022] [federalsadler@sohu.com]
# [Clicker] is licensed under Mulan PSL v2.
# You can use this software according to the terms and conditions of the Mulan PSL v2.
# You may obtain a copy of Mulan PSL v2 at:
# http://license.coscl.org.cn/MulanPSL2
# THIS SOFTWARE IS PROVIDED ON AN "AS IS" BASIS, WITHOUT WARRANTIES OF ANY KIND,
# EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO NON-INFRINGEMENT,
# MERCHANTABILITY OR FIT FOR A PARTICULAR PURPOSE.
# See the Mulan PSL v2 for more details.
import datetime
import sqlite3
import time
import keyboard
import mouse
import pyautogui
import pyperclip
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMessageBox
import threading
import sys
from setting import SettingsData

event = threading.Event()

COMMAND_TYPE_SIMULATE_CLICK = "模拟点击"
COMMAND_TYPE_CUSTOM = "自定义"


# 编写一个空的类
class MainWork:
    """主要工作类"""

    def __init__(self, file_path, main_window):
        # 终止和暂停标志
        self.start_state = True
        self.suspended = False
        # 文件路径和主窗体
        self.file_path = file_path
        self.main_window = main_window
        # 读取配置文件
        self.settings = SettingsData()
        self.settings.init()
        # 在窗体中显示循环次数
        self.number = 1

    # 读取数据库中的数据
    def extrate_data(self):
        """读取数据库中的数据"""
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select * from 命令')
        list_instructions = cursor.fetchall()
        print(list_instructions)
        con.close()
        return list_instructions

    def control_window(self):
        """控制窗体的显示与隐藏"""
        if self.main_window.checkBox_2.isChecked() and self.main_window.isHidden():
            self.main_window.show()
        elif self.main_window.checkBox_2.isChecked() and self.main_window.isVisible():
            self.main_window.hide()

    def star_work(self):
        """主要工作"""
        # 读取数据库中的数据
        list_instructions = self.extrate_data()
        # 获取设置参数并初始化
        self.settings.init()
        # 控制窗体的显示与隐藏
        self.control_window()
        # 开始执行主要操作
        try:
            if len(list_instructions) != 0:
                keyboard.hook(self.abc)
                if self.main_window.radioButton.isChecked():
                    # 在窗体中显示循环次数
                    self.number = 1
                    while True:
                        # 如果状态为True执行无限循环
                        self.execute_instructions(list_instructions)
                        if not self.start_state:
                            self.main_window.plainTextEdit.appendPlainText('结束任务')
                            # self.main_window.display_running_time('结束计时')
                            break
                        if self.suspended:
                            event.clear()
                            event.wait(86400)
                        self.number += 1
                        time.sleep(self.settings.time_sleep)
                    # 窗体显示
                elif self.main_window.radioButton_2.isChecked():
                    # self.main_window.display_running_time('开始计时')
                    number = 1
                    # 如果状态为有限次循环
                    repeat_number = self.main_window.spinBox.value()
                    # while number <= repeat_number and start_state:
                    while number <= repeat_number:
                        self.execute_instructions(list_instructions)
                        if not self.start_state:
                            self.control_window()
                            self.main_window.plainTextEdit.appendPlainText('结束任务')
                            # self.main_window.display_running_time('结束计时')
                            break
                        if self.suspended:
                            event.clear()
                            event.wait(86400)
                        number += 1
                        time.sleep(self.settings.time_sleep)
                    self.main_window.plainTextEdit.appendPlainText('结束任务')
                    # 窗体显示
                    self.control_window()
                elif not self.main_window.radioButton.isChecked() and not self.main_window.radioButton_2.isChecked():
                    QMessageBox.information(self.main_window, "提示", "请设置执行循环次数！")
        except:
            pass

    def execute_instructions(self, list_instructions):
        """执行接受到的操作指令"""
        # 设置进度条为0
        self.main_window.progressBar.setValue(0)
        # 读取指令
        for i in range(len(list_instructions)):
            # list_instructions=(id,图像名称，指令类型，参数1，参数2，重复次数)
            # list_instructions=(2, '0', '移动鼠标', '→', '100', 1)
            # list_instructions=(1, 'dd.png', '图像点击', '左键单击', '', 1)
            # 读取指令类型和重复次数
            cmd_type = list_instructions[i][2]
            re_try = list_instructions[i][5]
            # 设置进度条
            x = (i + 1) / len(list_instructions) * 100
            self.main_window.progressBar.setValue(int(x))
            # 设置一个容器，用于存储参数
            list_ins = []

            # 图像识别点击的事件
            if cmd_type == "图像点击":
                # 读取图像名称
                img = (self.file_path + "/" + list_instructions[i][1]).replace('/', '//')
                # 取重复次数
                re_try = list_instructions[i][5]
                # 是否跳过参数
                skip = list_instructions[i][4]
                if list_instructions[i][3] == '左键单击':
                    list_ins = [1, 'left', img, skip]
                elif list_instructions[i][3] == '左键双击':
                    list_ins = [2, 'left', img, skip]
                elif list_instructions[i][3] == '右键单击':
                    list_ins = [1, 'right', img, skip]
                elif list_instructions[i][3] == '右键双击':
                    list_ins = [2, 'right', img, skip]
                # 执行鼠标点击事件
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 屏幕坐标点击事件
            elif cmd_type == '坐标点击':
                # 取x,y坐标的值
                x = int(list_instructions[i][4].split('-')[0])
                y = int(list_instructions[i][4].split('-')[1])
                z = int(list_instructions[i][4].split('-')[2])
                # 调用鼠标点击事件（点击次数，按钮类型，图像名称）
                if list_instructions[i][3] == '左键单击':
                    list_ins = [1, 'left', x, y]
                elif list_instructions[i][3] == '左键双击':
                    list_ins = [2, 'left', x, y]
                elif list_instructions[i][3] == '右键单击':
                    list_ins = [1, 'right', x, y]
                elif list_instructions[i][3] == '右键双击':
                    list_ins = [2, 'right', x, y]
                elif list_instructions[i][3] == '左键（自定义次数）':
                    list_ins = [z, 'left', x, y]
                # 执行鼠标点击事件
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 等待的事件
            elif cmd_type == '等待':
                wait_type = list_instructions[i][3]
                if wait_type == '等待':
                    wait_time = int(list_instructions[i][4])
                    QApplication.processEvents()
                    self.main_window.plainTextEdit.appendPlainText('等待时长' + str(wait_time) + '秒')
                    self.stop_time(wait_time)
                elif wait_type == '等待到指定时间':
                    target_time = list_instructions[i][4].split('+')[0].replace('-', '/')
                    interval_time = list_instructions[i][4].split('+')[1]
                    now_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
                    # 将now_time转换为时间格式
                    now_time = datetime.datetime.strptime(now_time, '%Y/%m/%d %H:%M:%S')
                    # 将target_time转换为时间格式
                    t_time = datetime.datetime.strptime(target_time, '%Y/%m/%d %H:%M:%S')
                    if t_time > now_time:
                        year_target = int(t_time.strftime('%Y'))
                        month_target = int(t_time.strftime('%m'))
                        day_target = int(t_time.strftime('%d'))
                        hour_target = int(t_time.strftime('%H'))
                        minute_target = int(t_time.strftime('%M'))
                        second_target = int(t_time.strftime('%S'))
                        self.check_time(year_target, month_target, day_target,
                                        hour_target, minute_target, second_target, interval_time)

            # 滚轮滑动的事件
            elif cmd_type == '滚轮滑动':
                scroll_direction = list_instructions[i][3]
                scroll_distance = int(list_instructions[i][4])
                if scroll_direction == '↑':
                    scroll_distance = scroll_distance
                elif scroll_direction == '↓':
                    scroll_distance = -scroll_distance
                list_ins = [scroll_direction, scroll_distance]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 文本输入的事件
            elif cmd_type == '文本输入':
                input_value = str(list_instructions[i][3])
                list_ins = [input_value]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 鼠标移动的事件
            elif cmd_type == '鼠标移动':
                try:
                    direction = list_instructions[i][3]
                    distance = list_instructions[i][4]
                    list_ins = [direction, distance]
                    self.execution_repeats(cmd_type, list_ins, re_try)
                except IndexError:
                    self.main_window.plainTextEdit.appendPlainText('鼠标移动参数格式错误！')

            # 键盘按键的事件
            elif cmd_type == '按下键盘':
                key = list_instructions[i][3]
                list_ins = [key]
                self.execution_repeats(cmd_type, list_ins, re_try)
            # 中键激活的事件
            elif cmd_type == '中键激活':
                command_type = list_instructions[i][3]
                click_count = list_instructions[i][4]
                list_ins = [command_type, click_count]
                self.execution_repeats(cmd_type, list_ins, re_try)

    def execution_repeats(self, cmd_type, list_ins, reTry):
        """执行重复次数"""

        def determine_execution_type(cmd_type, list_ins):
            """执行判断命令类型并调用对应函数"""
            # 图像点击的操作事件
            if cmd_type == '图像点击':
                click_times = list_ins[0]
                lOrR = list_ins[1]
                img = list_ins[2]
                skip = list_ins[3]
                self.execute_click(click_times, lOrR, img, skip)
            # 坐标点击的操作事件
            elif cmd_type == '坐标点击':
                x = list_ins[2]
                y = list_ins[3]
                click_times = list_ins[0]
                lOrR = list_ins[1]
                # pyautogui.moveTo(x, y)
                pyautogui.click(x, y, click_times, interval=self.settings.interval, duration=self.settings.duration,
                                button=lOrR)
                self.main_window.plainTextEdit.appendPlainText('执行坐标%s:%s点击' % (x, y) + str(self.number))

            elif cmd_type == '鼠标移动':
                direction = list_ins[0]
                distance = list_ins[1]
                self.mouse_moves(direction, distance)
            elif cmd_type == '滚轮滑动':
                scroll_distance = list_ins[1]
                scroll_direction = list_ins[0]
                self.wheel_slip(scroll_direction, scroll_distance)
            elif cmd_type == '文本输入':
                input_value = list_ins[0]
                self.text_input(input_value)
            elif cmd_type == '按下键盘':
                # 获取键盘按键
                keys = list_ins[0].split('+')
                # 按下键盘
                if len(keys) == 1:
                    pyautogui.press(keys[0])  # 如果只有一个键,直接按下
                else:
                    # 否则,组合多个键为热键
                    hotkey = '+'.join(keys)
                    pyautogui.hotkey(hotkey)
                time.sleep(self.settings.time_sleep)
                self.main_window.plainTextEdit.appendPlainText('已经按下按键' + list_ins[0])
            elif cmd_type == '中键激活':
                command_type = list_ins[0]
                click_count = list_ins[1]
                self.middle_mouse_button(command_type, click_count)

        if reTry == 1:
            # 参数：图片和查找精度，返回目标图像在屏幕的位置
            determine_execution_type(cmd_type, list_ins)
        elif reTry > 1:
            # 有限次重复
            i = 1
            while i < reTry + 1:
                determine_execution_type(cmd_type, list_ins)
                i += 1
                time.sleep(self.settings.time_sleep)
        else:
            pass

    def check_time(self, year_target, month_target, day_target, hour_target, minute_target, second_target, inrerval):
        """检查时间，指定时间则执行操作"""
        show_times = 1
        sleep_time = int(inrerval) / 1000
        while True:
            now = time.localtime()
            if show_times == 1:
                QApplication.processEvents()
                self.main_window.plainTextEdit.appendPlainText(
                    '当前时间为：%s/%s/%s %s:%s:%s' % (now.tm_year, now.tm_mon,
                                                      now.tm_mday, now.tm_hour,
                                                      now.tm_min, now.tm_sec))
                print("当前时间为：%s/%s/%s %s:%s:%s" % (now.tm_year, now.tm_mon,
                                                        now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec))
                show_times = sleep_time
            if now.tm_year == year_target and now.tm_mon == month_target and \
                    now.tm_mday == day_target and now.tm_hour == hour_target and \
                    now.tm_min == minute_target and now.tm_sec == second_target:
                print("退出等待")
                self.main_window.plainTextEdit.appendPlainText('退出等待')
                break
            # 时间暂停
            time.sleep(sleep_time)
            show_times += sleep_time

    def middle_mouse_button(self, command_type, click_times):
        """中键点击事件"""
        QApplication.processEvents()
        self.main_window.plainTextEdit.appendPlainText('等待按下鼠标中键中...按下esc键退出')
        # 如果按下esc键则退出
        mouse.wait(button='middle')
        try:
            if command_type == COMMAND_TYPE_SIMULATE_CLICK:
                # print('执行鼠标点击'+click_times+'次')
                pyautogui.click(clicks=int(click_times), button='left')
                self.main_window.plainTextEdit.appendPlainText('执行鼠标点击' + click_times + '次')
            elif command_type == COMMAND_TYPE_CUSTOM:
                pass
        except OSError:
            # 弹出提示框。提示检查鼠标是否连接
            QMessageBox.critical(self.main_window, '提示', '连接失败，请检查鼠标是否连接正确。')
            pass

    def execute_click(self, click_times, lOrR, img, skip):
        """执行鼠标点击事件"""

        # 4个参数：鼠标点击时间，按钮类型（左键右键中键），图片名称，重复次数
        repeat = True
        number_1 = 1

        def image_match_click(remind):
            nonlocal repeat, number_1
            if location is not None:
                # 参数：位置X，位置Y，点击次数，时间间隔，持续时间，按键
                print('找到匹配图片' + str(self.number))
                pyautogui.click(location.x, location.y,
                                clicks=click_times, interval=self.settings.interval, duration=self.settings.duration,
                                button=lOrR)
                print('执行鼠标点击' + str(self.number))
                self.main_window.plainTextEdit.appendPlainText('执行鼠标点击' + str(self.number))
                self.real_time_display_status()
                repeat = False
            else:
                if remind:
                    self.main_window.plainTextEdit.appendPlainText(
                        '未匹配到图片' + str(self.number) + '正在重试' + str(number_1))
                    number_1 += 1
                else:
                    self.main_window.plainTextEdit.appendPlainText('未匹配到图片' + str(self.number))
                self.real_time_display_status()
                print('未找到匹配图片' + str(self.number))

        # location = pyautogui.locateCenterOnScreen(img, confidence=setting.confidence)
        try:
            print(img)
            if skip == "自动略过":
                print('执行自动略过')
                location = pyautogui.locateCenterOnScreen(img, confidence=self.settings.confidence)
                image_match_click(False)
            else:
                print('未执行自动略过')
                print(self.start_state)
                print(repeat)
                while self.start_state and repeat:
                    print('执行图像点击')
                    location = pyautogui.locateCenterOnScreen(img, confidence=self.settings.confidence)
                    print(location)
                    image_match_click(True)
                print("已完成执行图像点击")
        except OSError:
            QMessageBox.critical(self.main_window, '错误', '目标图像文件夹、图片命名或路径暂不支持中文！')

    def mouse_moves(self, direction, distance):
        """鼠标移动事件"""
        # 显示鼠标当前位置
        x, y = pyautogui.position()
        print('x:' + str(x) + ',y:' + str(y))
        # 相对于当前位置移动鼠标
        if direction == '↑':
            pyautogui.moveRel(0, -abs(int(distance)), duration=self.settings.duration)
        elif direction == '↓':
            pyautogui.moveRel(0, int(distance), duration=self.settings.duration)
        elif direction == '←':
            pyautogui.moveRel(-abs(int(distance)), 0, duration=self.settings.duration)
        elif direction == '→':
            pyautogui.moveRel(int(distance), 0, duration=self.settings.duration)
        self.main_window.plainTextEdit.appendPlainText('移动鼠标' + direction + distance + '像素距离')
        self.real_time_display_status()

    def wheel_slip(self, scroll_direction, scroll_distance):
        """滚轮滑动事件"""
        pyautogui.scroll(scroll_distance)
        self.main_window.plainTextEdit.appendPlainText(
            '滚轮滑动' + str(scroll_direction) + str(abs(scroll_distance)) + '距离')
        self.real_time_display_status()

    def text_input(self, input_value):
        """文本输入事件"""
        print(input_value)
        print("执行文本输入")
        pyperclip.copy(input_value)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(self.settings.time_sleep)
        self.main_window.plainTextEdit.appendPlainText('执行文本输入')

    def real_time_display_status(self):
        """设置实时显示状态文本"""
        QApplication.processEvents()
        # 当信息超过200行则清空
        self.main_window.clear_plaintext(200)

    def stop_time(self, seconds):
        """暂停时间"""
        for i in range(seconds):
            keyboard.hook(self.abc)
            QApplication.processEvents()
            # 显示剩下等待时间
            self.main_window.plainTextEdit.appendPlainText('等待中...剩余' + str(seconds - i) + '秒')
            if self.start_state is False:
                break
            # if self.suspended:
            #     # 显示暂停
            #     QApplication.processEvents()
            #     self.main_window.plainTextEdit.appendPlainText('暂停中...')
            #     event.clear()
            #     event.wait(86400)
            time.sleep(1)

    def abc(self, x):
        """键盘事件，退出任务、开始任务、暂停恢复任务"""
        a = keyboard.KeyboardEvent('down', 1, 'esc')
        s = keyboard.KeyboardEvent('down', 31, 's')
        r = keyboard.KeyboardEvent('down', 19, 'r')
        # var = x.scan_code
        # print(var)
        if x.event_type == 'down' and x.name == a.name:
            print("你按下了退出键")
            self.main_window.plainTextEdit.appendPlainText('你按下了退出键')
            self.start_state = False
        if x.event_type == 'down' and x.name == s.name:
            print("你按下了暂停键")
            self.main_window.plainTextEdit.appendPlainText('你按下了暂停键')
            self.suspended = True
        if x.event_type == 'down' and x.name == r.name:
            print('你按下了恢复键')
            self.main_window.plainTextEdit.appendPlainText('你按下了恢复键')
            self.suspended = False


def exit_main_work():
    sys.exit()