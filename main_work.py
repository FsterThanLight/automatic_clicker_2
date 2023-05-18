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
import os
import sqlite3
import sys
import threading
import time

import keyboard
import mouse
import openpyxl
import pyautogui
import pyperclip

event = threading.Event()

COMMAND_TYPE_SIMULATE_CLICK = "模拟点击"
COMMAND_TYPE_CUSTOM = "自定义"


# 编写一个空的类
class MainWork:
    """主要工作类"""

    def __init__(self, main_window):
        # 终止和暂停标志
        self.start_state = True
        self.suspended = False
        # 主窗体
        self.main_window = main_window
        # 读取配置文件
        self.settings = SettingsData()
        self.settings.init()
        # 在窗体中显示循环次数
        self.number = 1
        # 全部指令的循环次数，无限循环为标志
        self.infinite_cycle = False
        self.number_cycles = 1
        # 从数据库中读取全局参数
        self.image_folder_path, self.excel_folder_path, \
            self.branch_table_name, self.extenders = self.extracted_data_global_parameter()

    # def accdb(self):
    #     """建立与数据库的连接，返回游标"""
    #     try:
    #         path = os.path.abspath('.')
    #         # 取得当前文件目录
    #         mdb = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path + '\\' + self.odbc_name
    #         # 连接字符串
    #         conn = pypyodbc.win_connect_mdb(mdb)
    #         # 建立连接
    #         cursor = conn.cursor()
    #         print('成功连接数据库！')
    #         return cursor, conn
    #     except pypyodbc.Error:
    #         x = input("未连接到数据库！！请检查数据库路径是否异常。")
    #         sys.exit()

    def sqlitedb(self):
        """建立与数据库的连接，返回游标"""
        try:
            path = os.path.abspath('.')
            # 取得当前文件目录
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            print('成功连接数据库！')
            return cursor, con
        except sqlite3.Error:
            x = input("未连接到数据库！！请检查数据库路径是否异常。")
            sys.exit()

    def close_database(self, cursor, conn):
        """关闭数据库"""
        cursor.close()
        conn.close()

    def test(self):
        print('test')
        all_list_instructions = self.extracted_data_all_list(self.branch_table_name)
        print(all_list_instructions)
        print(len(all_list_instructions))

    def extracted_data_all_list(self, branch_table_name):
        """提取指令集中的数据,返回主表和分支表的汇总数据"""
        all_list_instructions = []
        # 从主表中提取数据
        cursor, conn = self.sqlitedb()
        cursor.execute("select * from 命令")
        main_list_instructions = cursor.fetchall()
        all_list_instructions.append(main_list_instructions)
        # 从分支表中提取数据
        if len(branch_table_name) != 0:
            for i in branch_table_name:
                cursor.execute("select * from " + i)
                branch_list_instructions = cursor.fetchall()
                all_list_instructions.append(branch_list_instructions)
        self.close_database(cursor, conn)
        return all_list_instructions

    # 编写一个函数用于去除列表中的none
    def remove_none(self, list_):
        """去除列表中的none"""
        list_x = []
        for i in list_:
            if i[0] is not None:
                list_x.append(i[0].replace('"', ''))
        return list_x

    def extracted_data_global_parameter(self):
        """从全局参数表中提取数据"""
        cursor, conn = self.sqlitedb()
        cursor.execute("select 图像文件夹路径 from 全局参数")
        image_folder_path = self.remove_none(cursor.fetchall())
        cursor.execute("select 工作簿路径 from 全局参数")
        excel_folder_path = self.remove_none(cursor.fetchall())
        cursor.execute("select 分支表名 from 全局参数")
        branch_table_name = self.remove_none(cursor.fetchall())
        cursor.execute("select 扩展程序 from 全局参数")
        extenders = self.remove_none(cursor.fetchall())
        self.close_database(cursor, conn)
        print("全局参数读取成功！")
        return image_folder_path, excel_folder_path, branch_table_name, extenders

    def start_work(self):
        """主要工作"""
        self.start_state = True
        self.suspended = False
        # 读取数据库中的数据
        list_instructions = self.extracted_data_all_list(self.branch_table_name)
        # 开始执行主要操作
        if len(list_instructions) != 0:
            keyboard.hook(self.abc)
            # # 如果状态为True执行无限循环
            if self.infinite_cycle:
                self.number = 1
                while True:
                    self.execute_instructions(0, 0, list_instructions)
                    if not self.start_state:
                        print('结束任务')
                        break
                    if self.suspended:
                        event.clear()
                        event.wait(86400)
                    self.number += 1
                    time.sleep(self.settings.time_sleep)
            # 如果状态为有限次循环
            elif self.infinite_cycle == False and self.number_cycles > 0:
                number = 1
                repeat_number = self.number_cycles
                while number <= repeat_number and self.start_state:
                    self.execute_instructions(0, 0, list_instructions)
                    if not self.start_state:
                        print('结束任务')
                        break
                    if self.suspended:
                        event.clear()
                        event.wait(86400)
                    number += 1
                    time.sleep(self.settings.time_sleep)
                print('结束任务')
            elif not self.infinite_cycle and self.number_cycles <= 0:
                print("请设置执行循环次数！")

    def execute_instructions(self, current_list_index, current_index, list_instructions):
        """执行接受到的操作指令"""
        # 读取指令
        while current_index < len(list_instructions[current_list_index]):
            elem = list_instructions[current_list_index][current_index]
            # 【指令集合【指令分支（指令元素[元素索引]）】】
            # print('执行当前指令：', elem)
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
            # 读取指令类型
            cmd_type = dict(dic)['指令类型']
            re_try = dict(dic)['重复次数']
            # 设置一个容器，用于存储参数
            list_ins = []

            # 图像识别点击的事件
            if cmd_type == "图像点击":
                # 读取图像名称
                img = list_instructions[current_list_index][1]
                # 取重复次数
                re_try = list_instructions[current_list_index][7]
                # 是否跳过参数
                skip = list_instructions[current_list_index][4]
                if list_instructions[current_list_index][3] == '左键单击':
                    list_ins = [1, 'left', img, skip]
                elif list_instructions[current_list_index][3] == '左键双击':
                    list_ins = [2, 'left', img, skip]
                elif list_instructions[current_list_index][3] == '右键单击':
                    list_ins = [1, 'right', img, skip]
                elif list_instructions[current_list_index][3] == '右键双击':
                    list_ins = [2, 'right', img, skip]
                # 执行鼠标点击事件
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 屏幕坐标点击事件
            elif cmd_type == '坐标点击':
                # 取x,y坐标的值
                x = int(dict(dic)['参数2'].split('-')[0])
                y = int(dict(dic)['参数2'].split('-')[1])
                z = int(dict(dic)['参数2'].split('-')[1])
                print('x,y坐标：', x, y)
                # 调用鼠标点击事件（点击次数，按钮类型，图像名称）
                if dict(dic)['参数1（键鼠指令）'] == '左键单击':
                    list_ins = [1, 'left', x, y]
                elif dict(dic)['参数1（键鼠指令）'] == '左键双击':
                    list_ins = [2, 'left', x, y]
                elif dict(dic)['参数1（键鼠指令）'] == '右键单击':
                    list_ins = [1, 'right', x, y]
                elif dict(dic)['参数1（键鼠指令）'] == '右键双击':
                    list_ins = [2, 'right', x, y]
                elif dict(dic)['参数1（键鼠指令）'] == '左键（自定义次数）':
                    list_ins = [z, 'left', x, y]
                # 执行鼠标点击事件
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 等待的事件
            elif cmd_type == '等待':
                wait_type = list_instructions[current_list_index][3]
                if wait_type == '等待':
                    wait_time = int(list_instructions[current_list_index][4])
                    print('等待时长' + str(wait_time) + '秒')
                    self.stop_time(wait_time)
                elif wait_type == '等待到指定时间':
                    target_time = list_instructions[current_list_index][4].split('+')[0].replace('-', '/')
                    interval_time = list_instructions[current_list_index][4].split('+')[1]
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
                scroll_direction = list_instructions[current_list_index][3]
                scroll_distance = int(list_instructions[current_list_index][4])
                if scroll_direction == '↑':
                    scroll_distance = scroll_distance
                elif scroll_direction == '↓':
                    scroll_distance = -scroll_distance
                list_ins = [scroll_direction, scroll_distance]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 文本输入的事件
            elif cmd_type == '文本输入':
                input_value = str(list_instructions[current_list_index][3])
                list_ins = [input_value]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 鼠标移动的事件
            elif cmd_type == '鼠标移动':
                try:
                    direction = list_instructions[current_list_index][3]
                    distance = list_instructions[current_list_index][4]
                    list_ins = [direction, distance]
                    self.execution_repeats(cmd_type, list_ins, re_try)
                except IndexError:
                    print('鼠标移动参数格式错误！')

            # 键盘按键的事件
            elif cmd_type == '按下键盘':
                key = list_instructions[current_list_index][3]
                list_ins = [key]
                self.execution_repeats(cmd_type, list_ins, re_try)
            # 中键激活的事件
            elif cmd_type == '中键激活':
                command_type = list_instructions[current_list_index][3]
                click_count = list_instructions[current_list_index][4]
                list_ins = [command_type, click_count]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 鼠标事件
            elif cmd_type == '鼠标事件':
                if list_instructions[current_list_index][3] == '左键单击':
                    list_ins = [1, 'left']
                elif list_instructions[current_list_index][3] == '左键双击':
                    list_ins = [2, 'left']
                elif list_instructions[current_list_index][3] == '右键单击':
                    list_ins = [1, 'right']
                elif list_instructions[current_list_index][3] == '右键双击':
                    list_ins = [2, 'right']
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 图片信息录取
            elif cmd_type == '图像信息录入':
                excel_path = list_instructions[current_list_index][4].replace('"', '')
                img = list_instructions[current_list_index][1].replace('"', '')
                cell_position = list_instructions[current_list_index][5]
                exception_type = list_instructions[current_list_index][7]
                list_ins = [3, 'left', img, excel_path, cell_position, exception_type]
                self.execution_repeats(cmd_type, list_ins, re_try)

            # 跳转分支的指定指令
            if elem[-1] == '':
                current_index += 1
            elif elem[-1] != '':
                branch_name_index, branch_index = elem[-1].split('-')
                x = int(branch_name_index) - 1
                y = int(branch_index) - 1
                self.execute_instructions(x, y, list_instructions)
                break

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
                pyautogui.click(x, y, click_times, interval=self.settings.interval, duration=self.settings.duration,
                                button=lOrR)
                print('执行坐标%s:%s点击' % (x, y) + str(self.number))

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
                print('已经按下按键' + list_ins[0])
            elif cmd_type == '中键激活':
                command_type = list_ins[0]
                click_count = list_ins[1]
                self.middle_mouse_button(command_type, click_count)
            elif cmd_type == '鼠标事件':
                click_times = list_ins[0]
                lOrR = list_ins[1]
                position = pyautogui.position()
                pyautogui.click(position[0], position[1], click_times, interval=self.settings.interval,
                                duration=self.settings.duration,
                                button=lOrR)
                print('执行鼠标事件')
            elif cmd_type == '图像信息录入':
                # 图像参数
                img = list_ins[2]
                # excel参数
                excel_path = list_ins[3]
                cell_position = list_ins[4]
                # 鼠标单击参数
                click_times = list_ins[0]
                lOrR = list_ins[1]
                exception_type = list_ins[5]
                # 获取excel表格中的值
                cell_value = self.extra_excel_cell_value(excel_path, cell_position)
                self.execute_click(click_times, lOrR, img, exception_type)
                self.text_input(cell_value)
                print('已执行信息录入')

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

    def extra_excel_cell_value(self, excel_path, sheet_name, cell_position):
        """获取excel表格中的值"""
        try:
            # 打开excel表格
            wb = openpyxl.load_workbook(excel_path)
            # 选择表格
            sheet = wb[str(sheet_name)]
            # 获取单元格的值
            cell_value = sheet[cell_position].value
            print('获取到的单元格值为：' + str(cell_value))
            return cell_value
        except FileNotFoundError:
            x = input('没有找到工作簿')
            exit_main_work()
        except KeyError:
            x = input('没有找到表格')
            exit_main_work()
        except AttributeError:
            x = input('单元格格式错误')
            exit_main_work()

    def check_time(self, year_target, month_target, day_target, hour_target, minute_target, second_target, inrerval):
        """检查时间，指定时间则执行操作"""
        show_times = 1
        sleep_time = int(inrerval) / 1000
        while True:
            now = time.localtime()
            if show_times == 1:
                print("当前时间为：%s/%s/%s %s:%s:%s" % (now.tm_year, now.tm_mon,
                                                        now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec))
                show_times = sleep_time
            if now.tm_year == year_target and now.tm_mon == month_target and \
                    now.tm_mday == day_target and now.tm_hour == hour_target and \
                    now.tm_min == minute_target and now.tm_sec == second_target:
                print("退出等待")
                break
            # 时间暂停
            time.sleep(sleep_time)
            show_times += sleep_time

    def middle_mouse_button(self, command_type, click_times):
        """中键点击事件"""
        print('等待按下鼠标中键中...按下esc键退出')
        # 如果按下esc键则退出
        mouse.wait(button='middle')
        try:
            if command_type == COMMAND_TYPE_SIMULATE_CLICK:
                # print('执行鼠标点击'+click_times+'次')
                pyautogui.click(clicks=int(click_times), button='left')
                print('执行鼠标点击' + click_times + '次')
            elif command_type == COMMAND_TYPE_CUSTOM:
                pass
        except OSError:
            # 弹出提示框。提示检查鼠标是否连接
            print('连接失败，请检查鼠标是否连接正确。')
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
                # self.real_time_display_status()
                repeat = False
            else:
                if remind:
                    print('未找到匹配图片' + str(self.number) + '正在重试' + str(number_1))
                    number_1 += 1
                else:
                    print('未找到匹配图片' + str(self.number))
                # self.real_time_display_status()

        # location = pyautogui.locateCenterOnScreen(img, confidence=setting.confidence)
        try:
            print(img)
            if skip == "自动略过":
                print('执行自动略过')
                location = pyautogui.locateCenterOnScreen(img, confidence=self.settings.confidence)
                image_match_click(False)
            else:
                while self.start_state and repeat:
                    print('执行图像点击')
                    location = pyautogui.locateCenterOnScreen(img, confidence=self.settings.confidence)
                    print(location)
                    image_match_click(True)
        except OSError:
            print('目标图像文件夹、图片命名或路径暂不支持中文！')

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
        print('移动鼠标' + direction + distance + '像素距离')
        # self.real_time_display_status()

    def wheel_slip(self, scroll_direction, scroll_distance):
        """滚轮滑动事件"""
        pyautogui.scroll(scroll_distance)
        print('滚轮滑动' + str(scroll_direction) + str(abs(scroll_distance)) + '距离')

    def text_input(self, input_value):
        """文本输入事件"""
        pyperclip.copy(input_value)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(self.settings.time_sleep)
        print('执行文本输入' + str(input_value))

    def stop_time(self, seconds):
        """暂停时间"""
        for i in range(seconds):
            keyboard.hook(self.abc)
            # 显示剩下等待时间
            print('等待中...剩余' + str(seconds - i) + '秒')
            if self.start_state is False:
                break
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
            self.start_state = False
        if x.event_type == 'down' and x.name == s.name:
            print("你按下了暂停键")
            self.suspended = True
        if x.event_type == 'down' and x.name == r.name:
            print('你按下了恢复键')
            self.suspended = False


class SettingsData:
    def __init__(self):
        self.duration = 0
        self.interval = 0
        self.confidence = 0
        self.time_sleep = 0

    def init(self):
        """设置初始化"""
        # 从数据库加载设置
        """建立与数据库的连接，返回游标"""
        # 取得当前文件目录
        cursor, conn = self.sqlitedb()
        # 从数据库中取出全部数据
        cursor.execute('select * from 设置')
        # 读取全部数据
        list_setting_data = cursor.fetchall()
        # 关闭连接
        self.close_database(cursor, conn)
        # print(list_setting_data)

        for i in range(len(list_setting_data)):
            if list_setting_data[i][0] == '图像匹配精度':
                self.confidence = list_setting_data[i][1]
            elif list_setting_data[i][0] == '时间间隔':
                self.interval = list_setting_data[i][1]
            elif list_setting_data[i][0] == '持续时间':
                self.duration = list_setting_data[i][1]
            elif list_setting_data[i][0] == '暂停时间':
                self.time_sleep = list_setting_data[i][1]

    def sqlitedb(self):
        """建立与数据库的连接，返回游标"""
        try:
            # path = os.path.abspath('.')
            # # 取得当前文件目录
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            print('成功连接数据库！')
            return cursor, con
        except sqlite3.Error:
            x = input("未连接到数据库！！请检查数据库路径是否异常。")
            sys.exit()

    def close_database(self, cursor, conn):
        """关闭数据库"""
        cursor.close()
        conn.close()


def exit_main_work():
    sys.exit()

# if __name__ == '__main__':
#     # x = input('按回车键开始')
#     # odbc_name = '命令集.accdb'
#     # main_work = MainWork(odbc_name)
#     # main_work.start_work()
#     # y = input('按回车键退出')
#
#     # test
#     odbc_name = '命令集.accdb'
#     main_work = MainWork(odbc_name)
#     main_work.test()
