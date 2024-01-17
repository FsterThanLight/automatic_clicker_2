# coding: utf-8
# Copyright (c) [2022] [federalsadler@sohu.com]
# [Clicker] is licensed under Mulan PSL v2.
# You can use this software according to the terms and conditions of the Mulan PSL v2.
# You may obtain a copy of Mulan PSL v2 at:
# http://license.coscl.org.cn/MulanPSL2
# THIS SOFTWARE IS PROVIDED ON AN "AS IS" BASIS, WITHOUT WARRANTIES OF ANY KIND,
# EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO NON-INFRINGEMENT,
# MERCHANTABILITY OR FIT FOR A PARTICULAR PURPOSE.
# See the Mulan PSL v2 for more details.
import sqlite3
import subprocess

#
import winsound

from 功能类 import *
from 数据库操作 import sqlitedb, close_database, extract_global_parameter
from 网页操作 import WebOption

COMMAND_TYPE_SIMULATE_CLICK = "模拟点击"
COMMAND_TYPE_CUSTOM = "自定义"


def exit_main_work():
    sys.exit()


class MainWork:
    """主要工作类"""

    def __init__(self, main_window, navigation):
        # 终止和暂停标志
        self.start_state = True
        self.suspended = False
        # 主窗体
        self.main_window = main_window
        # 导航窗体
        self.navigation = navigation
        # 网页操作类
        self.web_option = WebOption(self.main_window, self.navigation)
        # 读取配置文件
        self.settings = SettingsData()
        self.settings.init()
        # 在窗体中显示循环次数
        self.number = 1
        # 全部指令的循环次数，无限循环为标志
        self.infinite_cycle = self.main_window.radioButton.isChecked()
        self.number_cycles = self.main_window.spinBox.value()
        # 从数据库中读取全局参数
        self.image_folder_path = extract_global_parameter('资源文件夹路径')
        self.branch_table_name = extract_global_parameter('分支表名')

    def reset_loop_count_and_infinite_loop_judgment(self):
        """重置循环次数和无限循环标志"""
        self.number = 1
        self.infinite_cycle = self.main_window.radioButton.isChecked()
        self.number_cycles = self.main_window.spinBox.value()

    def extracted_data_all_list(self, only_current_instructions=False) -> list:
        """提取指令集中的数据,返回主表和分支表的汇总数据
        :param only_current_instructions: 是否只提取当前分支的指令"""
        all_list_instructions = []
        # 从主表中提取数据
        cursor, conn = sqlitedb()
        # 从分支表中提取数据
        try:
            if not only_current_instructions:
                if len(self.branch_table_name) != 0:
                    for i in self.branch_table_name:
                        cursor.execute("select * from 命令 where 隶属分支 = ?", (i,))
                        branch_list_instructions = cursor.fetchall()
                        all_list_instructions.append(branch_list_instructions)
            if only_current_instructions:
                cursor.execute("select * from 命令 where 隶属分支 = ?", (self.main_window.comboBox.currentText(),))
                branch_list_instructions = cursor.fetchall()
                all_list_instructions.append(branch_list_instructions)
            close_database(cursor, conn)
            return all_list_instructions
        except sqlite3.OperationalError:
            QMessageBox.critical(self.main_window, "警告", "找不到分支！请检查分支表名是否正确！", QMessageBox.Yes)

    def start_work(self, only_current_instructions=False):
        """主要工作"""
        self.start_state = True
        self.suspended = False
        # 打印循环次数
        self.reset_loop_count_and_infinite_loop_judgment()
        # 读取数据库中的数据
        list_instructions = self.extracted_data_all_list(only_current_instructions)
        # 开始执行主要操作
        self.main_window.plainTextEdit.clear()
        self.main_window.tabWidget.setCurrentIndex(0)
        try:
            if len(list_instructions) != 0:
                # keyboard.hook(self.abc)
                # # 如果状态为True执行无限循环
                if self.infinite_cycle:
                    self.number = 1
                    while self.start_state:
                        self.execute_instructions(0, 0, list_instructions)
                        if not self.start_state:
                            self.main_window.plainTextEdit.appendPlainText('结束任务')
                            break
                        if self.suspended:
                            pass
                            # event.clear()
                            # event.wait(86400)
                        self.number += 1
                        time.sleep(self.settings.time_sleep)
                # 如果状态为有限次循环
                elif not self.infinite_cycle and self.number_cycles > 0:
                    self.number = 1
                    repeat_number = self.number_cycles
                    while self.number <= repeat_number and self.start_state:
                        self.execute_instructions(0, 0, list_instructions)
                        if not self.start_state:
                            self.main_window.plainTextEdit.appendPlainText('结束任务')
                            break
                        if self.suspended:
                            pass
                            # event.clear()
                            # event.wait(86400)
                        # print('第', self.number, '次循环')
                        self.main_window.plainTextEdit.appendPlainText('完成第' + str(self.number) + '次循环')
                        self.number += 1
                        time.sleep(self.settings.time_sleep)
                    self.main_window.plainTextEdit.appendPlainText('结束任务')
                elif not self.infinite_cycle and self.number_cycles <= 0:
                    print("请设置执行循环次数！")
        finally:
            self.web_option.close_browser()

    def execute_instructions(self, current_list_index, current_index, list_instructions):
        """执行接受到的操作指令"""
        # 读取指令
        while current_index < len(list_instructions[current_list_index]):
            try:
                elem_ = list_instructions[current_list_index][current_index]
                print('elem_:', elem_)
                # 【指令集合【指令分支（指令元素[元素索引]）】】
                # print('执行当前指令：', elem_)
                dic_ = {
                    'ID': elem_[0],
                    '图像路径': elem_[1],
                    '指令类型': elem_[2],
                    '参数1（键鼠指令）': elem_[3],
                    '参数2': elem_[4],
                    '参数3': elem_[5],
                    '参数4': elem_[6],
                    '重复次数': elem_[7],
                    '异常处理': elem_[8]
                }
                # 读取指令类型
                cmd_type = dict(dic_)['指令类型']
                exception_handling = dict(dic_)['异常处理']
                try:
                    # 图像识别点击的事件
                    if cmd_type == "图像点击":
                        image_click = ImageClick(main_window=self.main_window, ins_dic=dic_)
                        image_click.start_execute(self.number)

                    # 屏幕坐标点击事件
                    elif cmd_type == '坐标点击':
                        coordinate_click = CoordinateClick(main_window=self.main_window, ins_dic=dic_)
                        coordinate_click.start_execute(self.number)

                    # 等待的事件
                    elif cmd_type == '时间等待':
                        waiting = TimeWaiting(main_window=self.main_window, ins_dic=dic_)
                        waiting.start_execute()

                    # 图像等待事件
                    elif cmd_type == '图像等待':
                        image_waiting = ImageWaiting(main_window=self.main_window, ins_dic=dic_)
                        image_waiting.start_execute()

                    # 滚轮滑动的事件
                    elif cmd_type == '滚轮滑动':
                        scroll_wheel = RollerSlide(main_window=self.main_window, ins_dic=dic_)
                        scroll_wheel.start_execute()

                    # 文本输入的事件
                    elif cmd_type == '文本输入':
                        text_input = TextInput(main_window=self.main_window, ins_dic=dic_)
                        text_input.start_execute()

                    # 鼠标移动的事件
                    elif cmd_type == '移动鼠标':
                        move_mouse = MoveMouse(main_window=self.main_window, ins_dic=dic_)
                        move_mouse.start_execute()

                    # 键盘按键的事件
                    elif cmd_type == '按下键盘':
                        press_keyboard = PressKeyboard(main_window=self.main_window, ins_dic=dic_)
                        press_keyboard.start_execute()

                    # 中键激活的事件
                    elif cmd_type == '中键激活':
                        middle_activation = MiddleActivation(main_window=self.main_window, ins_dic=dic_)
                        middle_activation.start_execute()

                    # 鼠标事件
                    elif cmd_type == '鼠标点击':
                        mouse_click = MouseClick(main_window=self.main_window, ins_dic=dic_)
                        mouse_click.start_execute()

                    # 图片信息录取
                    elif cmd_type == '信息录入':
                        information_entry = InformationEntry(main_window=self.main_window, ins_dic=dic_)
                        information_entry.start_execute(self.number)

                    # 网页操作
                    elif cmd_type == '网页控制':
                        web_control = WebControl(main_window=self.main_window, ins_dic=dic_, navigation=self.navigation)
                        web_control.start_execute()

                    # 网页录入
                    elif cmd_type == '网页录入':
                        web_entry = WebEntry(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        web_entry.start_execute(self.number)

                    # 鼠标拖拽
                    elif cmd_type == '鼠标拖拽':
                        mouse_drag = MouseDrag(main_window=self.main_window, ins_dic=dic_)
                        mouse_drag.start_execute()

                    # 切换frame
                    elif cmd_type == '切换frame':
                        toggle_frame = ToggleFrame(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        toggle_frame.start_execute()

                    # 读取网页数据到excel
                    elif cmd_type == '保存表格':
                        save_form = SaveForm(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        save_form.start_execute()

                    # 拖动网页元素
                    elif cmd_type == '拖动元素':
                        drag_element = DragWebElements(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        drag_element.start_execute()

                    # 全屏截图
                    elif cmd_type == '全屏截图':
                        full_screen_capture = FullScreenCapture(main_window=self.main_window, ins_dic=dic_)
                        full_screen_capture.start_execute()

                    # 窗口切换
                    elif cmd_type == '切换窗口':
                        # 切换窗口
                        switch_window = SwitchWindow(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        switch_window.start_execute()

                    # 发送消息到微信
                    elif cmd_type == '发送消息':
                        sendwechat = SendWeChat(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        sendwechat.start_execute()

                    # 执行完毕后，跳转到下一条指令
                    current_index += 1
                except Exception as e:
                    # except AttributeError:
                    print(e)
                    # 跳转分支的指定指令
                    print('分支指令:' + exception_handling)

                    # 自动跳过功能
                    if exception_handling == '自动跳过':
                        current_index += 1
                    elif exception_handling == '抛出异常并暂停':
                        winsound.Beep(1000, 1000)
                        # 弹出提示框
                        reply = QMessageBox.question(self.main_window, '提示',
                                                     'ID为{}的指令抛出异常！\n是否继续执行？'.format(dict(dic_)['ID']),
                                                     QMessageBox.Yes | QMessageBox.No,
                                                     QMessageBox.No)
                        if reply == QMessageBox.Yes:
                            current_index += 1
                        else:
                            self.start_state = False
                            current_index += 1
                            break

                    # 抛出异常并停止
                    elif exception_handling == '抛出异常并停止':
                        winsound.Beep(1000, 1000)
                        # 弹出提示框
                        QMessageBox.warning(self.main_window, '提示',
                                            'ID为{}的指令抛出异常！\n已停止执行！'.format(dict(dic_)['ID']))
                        current_index += 1
                        self.start_state = False
                        break

                    # 使用扩展程序
                    elif exception_handling.endswith('.py') or exception_handling.endswith('.exe'):
                        self.start_state = False
                        self.main_window.plainTextEdit.appendPlainText('执行扩展程序')
                        if '.exe' in exception_handling:
                            subprocess.run('calc.exe')
                        elif '.py' in exception_handling:
                            subprocess.run('python {}'.format(exception_handling))
                        break

                    # 跳转分支指令
                    elif '分支' in exception_handling:  # 跳转分支
                        self.main_window.plainTextEdit.appendPlainText('转到分支')
                        branch_name_index = exception_handling.split('-')[1]
                        branch_index = exception_handling.split('-')[2]
                        x = int(branch_name_index)
                        y = int(branch_index)
                        print('x:', x, 'y:', y)
                        self.execute_instructions(x, y, list_instructions)
                        break

            except IndexError:
                self.main_window.plainTextEdit.appendPlainText('分支执行异常！')
                QMessageBox.warning(self.main_window, '提示', '分支执行异常！')
                exit_main_work()

    def abc(self, x):
        """键盘事件，退出任务、开始任务、暂停恢复任务"""
        pass
        # a = keyboard.KeyboardEvent('down', 1, 'esc')
        # s = keyboard.KeyboardEvent('down', 31, 's')
        # r = keyboard.KeyboardEvent('down', 19, 'r')
        # # var = x.scan_code
        # # print(var)
        # if x.event_type == 'down' and x.name == a.name:
        #     self.main_window.plainTextEdit.appendPlainText('你按下了退出键')
        #     print("你按下了退出键")
        #     self.start_state = False
        # if x.event_type == 'down' and x.name == s.name:
        #     self.main_window.plainTextEdit.appendPlainText('你按下了暂停键')
        #     print("你按下了暂停键")
        #     self.suspended = True
        # if x.event_type == 'down' and x.name == r.name:
        #     self.main_window.plainTextEdit.appendPlainText('你按下了恢复键')
        #     print('你按下了恢复键')
        #     self.suspended = False


class SettingsData:
    def __init__(self):
        self.duration = 0
        self.interval = 0
        self.confidence = 0
        self.time_sleep = 0

    def init(self):
        """设置初始化"""
        # 从数据库加载设置
        # 取得当前文件目录
        cursor, conn = sqlitedb()
        # 从数据库中取出全部数据
        cursor.execute('select * from 设置')
        # 读取全部数据
        list_setting_data = cursor.fetchall()
        # 关闭连接
        close_database(cursor, conn)

        for i in range(len(list_setting_data)):
            if list_setting_data[i][0] == '图像匹配精度':
                self.confidence = list_setting_data[i][1]
            elif list_setting_data[i][0] == '时间间隔':
                self.interval = list_setting_data[i][1]
            elif list_setting_data[i][0] == '持续时间':
                self.duration = list_setting_data[i][1]
            elif list_setting_data[i][0] == '暂停时间':
                self.time_sleep = list_setting_data[i][1]


if __name__ == '__main__':
    pass
