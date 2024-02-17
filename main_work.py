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

import pymsgbox
from PyQt5.QtCore import *

from 功能类 import *
from 数据库操作 import extract_global_parameter, extracted_ins_from_database, get_str_now_time, system_prompt_tone


class CommandThread(QThread):
    """指令线程"""
    send_message = pyqtSignal(str, name='send_message')
    finished_signal = pyqtSignal(str, name='finished_signal')
    send_type_and_id = pyqtSignal(str, str, name='send_type_and_id')

    def __init__(self, main_window, navigation):
        super(CommandThread, self).__init__(parent=None)
        # 窗体属性
        self.main_window = main_window
        self.navigation = navigation
        self.out_mes = OutputMessage(self, self.navigation)
        # 循环控制
        self.number: int = 1  # 在窗体中显示循环次数
        self.number_cycles: int = 0  # 循环次数
        # 终止和暂停标志
        self.start_state: bool = True
        self.suspended: bool = False
        # 运行时的参数
        self.branch_name_index: int = 0  # 分支表名索引
        # 读取配置文件
        self.time_sleep = float(get_setting_data_from_db('暂停时间'))
        self.image_folder_path = extract_global_parameter('资源文件夹路径')
        self.branch_table_name = extract_global_parameter('分支表名')
        # 互斥锁,用于暂停线程
        self.mutex = QMutex()
        self.condition = QWaitCondition()
        self.is_paused: bool = False

    def set_branch_name_index(self, branch_name_index_):
        """设置分支表名索引"""
        self.branch_name_index = branch_name_index_

    def show_message(self, message):
        """显示消息"""
        self.send_message.emit(message)

    def run(self):
        """执行指令"""
        self.start_state = True
        self.suspended = False
        # 执行指令
        list_instructions = extracted_ins_from_database()
        if len(list_instructions) != 0:
            # 设置主流程循环前的参数
            loop_type = '无限循环' if self.main_window.radioButton.isChecked() else '有限循环'
            self.number = 1
            self.number_cycles = int(self.main_window.spinBox.value())
            # 开始循环执行指令
            while (self.start_state and loop_type == '无限循环') or \
                    (loop_type == '有限循环' and self.number <= self.number_cycles):
                # 执行指令集中的指令
                self.execute_instructions(self.branch_name_index, 0, list_instructions)
                self.show_message(f'完成第{self.number}次循环')
                self.number += 1
                time.sleep(self.time_sleep)

            # 结束信号
            self.finished_signal.emit('任务完成')

    def pause(self):
        self.mutex.lock()
        self.is_paused = True
        self.mutex.unlock()
        print('暂停线程')

    def resume(self):
        self.mutex.lock()
        self.is_paused = False
        self.condition.wakeAll()
        self.mutex.unlock()
        print('恢复线程')

    def check_mutex(self):
        self.mutex.lock()
        while self.is_paused:
            self.condition.wait(self.mutex)
        self.mutex.unlock()

    def execute_instructions(self, current_list_index, current_index, list_instructions_):
        """执行接受到的操作指令"""
        # 读取指令
        while current_index < len(list_instructions_[current_list_index]) and not self.check_mutex():
            # while current_index < len(list_instructions_[current_list_index]):
            try:
                elem_ = list_instructions_[current_list_index][current_index]
                # print('elem_:', elem_)
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
                        image_click = ImageClick(outputmessage=self.out_mes, ins_dic=dic_)
                        image_click.start_execute(self.number)

                    # 屏幕坐标点击事件
                    elif cmd_type == '坐标点击':
                        coordinate_click = CoordinateClick(command_thread=self, ins_dic=dic_)
                        coordinate_click.start_execute(self.number)

                    # 等待的事件
                    elif cmd_type == '时间等待':
                        waiting = TimeWaiting(command_thread=self, ins_dic=dic_)
                        waiting.start_execute()

                    # 图像等待事件
                    elif cmd_type == '图像等待':
                        image_waiting = ImageWaiting(command_thread=self, ins_dic=dic_)
                        image_waiting.start_execute()

                    # 滚轮滑动的事件
                    elif cmd_type == '滚轮滑动':
                        scroll_wheel = RollerSlide(command_thread=self, ins_dic=dic_)
                        scroll_wheel.start_execute()

                    # 文本输入的事件
                    elif cmd_type == '文本输入':
                        text_input = TextInput(command_thread=self, ins_dic=dic_)
                        text_input.start_execute()

                    # 鼠标移动的事件
                    elif cmd_type == '移动鼠标':
                        move_mouse = MoveMouse(command_thread=self, ins_dic=dic_)
                        move_mouse.start_execute()

                    # 键盘按键的事件
                    elif cmd_type == '按下键盘':
                        press_keyboard = PressKeyboard(command_thread=self, ins_dic=dic_)
                        press_keyboard.start_execute()

                    # 中键激活的事件
                    elif cmd_type == '中键激活':
                        middle_activation = MiddleActivation(command_thread=self, ins_dic=dic_)
                        middle_activation.start_execute()

                    # 鼠标事件
                    elif cmd_type == '鼠标点击':
                        mouse_click = MouseClick(command_thread=self, ins_dic=dic_)
                        mouse_click.start_execute()

                    # 图片信息录取
                    elif cmd_type == '信息录入':
                        information_entry = InformationEntry(command_thread=self, ins_dic=dic_)
                        information_entry.start_execute(self.number)

                    # 网页操作
                    elif cmd_type == '打开网址':
                        web_control = OpenWeb(command_thread=self, ins_dic=dic_, navigation=self.navigation)
                        web_control.start_execute()

                    # 网页元素操作
                    elif cmd_type == '元素控制':
                        web_element = EleControl(command_thread=self, ins_dic=dic_, navigation=self.navigation)
                        web_element.start_execute()

                    # 网页录入
                    elif cmd_type == '网页录入':
                        web_entry = WebEntry(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        web_entry.start_execute(self.number)

                    # 鼠标拖拽
                    elif cmd_type == '鼠标拖拽':
                        mouse_drag = MouseDrag(command_thread=self, ins_dic=dic_)
                        mouse_drag.start_execute()

                    # 切换frame
                    elif cmd_type == '切换frame':
                        toggle_frame = ToggleFrame(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        toggle_frame.start_execute()

                    # 读取网页数据到excel
                    elif cmd_type == '保存表格':
                        save_form = SaveForm(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        save_form.start_execute()

                    # 拖动网页元素
                    elif cmd_type == '拖动元素':
                        drag_element = DragWebElements(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        drag_element.start_execute()

                    # 全屏截图
                    elif cmd_type == '全屏截图':
                        full_screen_capture = FullScreenCapture(command_thread=self, ins_dic=dic_)
                        full_screen_capture.start_execute()

                    # 窗口切换
                    elif cmd_type == '切换窗口':
                        # 切换窗口
                        switch_window = SwitchWindow(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        switch_window.start_execute()

                    # 发送消息到微信
                    elif cmd_type == '发送消息':
                        sendwechat = SendWeChat(
                            command_thread=self,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        sendwechat.start_execute()

                    # 数字验证码
                    elif cmd_type == '数字验证码':
                        digital_verification_code = VerificationCode(
                            main_window=self.main_window,
                            ins_dic=dic_,
                            navigation=self.navigation
                        )
                        digital_verification_code.start_execute()

                    # 执行完毕后，跳转到下一条指令
                    current_index += 1

                except Exception as e:
                    str_id = str(dict(dic_)['ID'])
                    # except IndexError:

                    # 自动跳过功能
                    if exception_handling == '自动跳过':
                        self.show_message(f'ID为{str_id}的指令执行异常，已自动跳过。')
                        current_index += 1

                    # 提示异常并暂停
                    elif exception_handling == '提示异常并暂停':
                        system_prompt_tone('执行异常')
                        self.show_message(f'ID为{str_id}的指令执行异常，已提示异常并暂停。')
                        # 弹出带有OK按钮的提示框
                        choice = pymsgbox.confirm(
                            text=f'ID为{str_id}的指令执行异常！\n是否重试？\n\n错误类型：{str(type(e))}',
                            title='提示',
                            buttons=[pymsgbox.ABORT_TEXT, pymsgbox.RETRY_TEXT, pymsgbox.IGNORE_TEXT])
                        # 选择的按钮
                        if choice == pymsgbox.RETRY_TEXT:  # 重试指令
                            pass
                        elif choice == pymsgbox.IGNORE_TEXT:  # 忽略该指令,继续执行下一条指令
                            current_index += 1
                        elif choice == pymsgbox.ABORT_TEXT:  # 终止任务
                            self.start_state = False
                            break

                    # 抛出异常并停止
                    elif exception_handling == '提示异常并停止':
                        system_prompt_tone('执行异常')
                        self.show_message(f'ID为{str_id}的指令执行异常，已提示异常并停止。')
                        # 弹出提示框
                        pymsgbox.alert(
                            text=f'ID为{str_id}的指令抛出异常！\n\n错误类型：{str(type(e))}',
                            title='提示',
                            icon=pymsgbox.STOP
                        )
                        current_index += 1
                        self.start_state = False
                        break

                    # 跳转分支指令
                    else:  # 跳转分支
                        self.show_message(f'转到分支：{exception_handling}')
                        target_branch_name = exception_handling.split('-')[0]  # 分支表名
                        # 目标分支表名在分支表名中的索引
                        branch_table_name_index = self.branch_table_name.index(target_branch_name)
                        # 分支表中要跳转的指令索引
                        branch_ins_index = exception_handling.split('-')[1]
                        x = int(branch_table_name_index)
                        y = int(branch_ins_index) - 1
                        self.execute_instructions(x, y, list_instructions_)
                        break

            except IndexError:
                self.show_message(f'分支执行异常！')
                break
