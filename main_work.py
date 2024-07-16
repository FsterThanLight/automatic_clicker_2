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

from PyQt5.QtCore import *

from functions import system_prompt_tone
from ini控制 import get_branch_info
from 功能类 import *
from 数据库操作 import extracted_ins_from_database, extracted_ins_target_id_from_database


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
        self.run_mode: tuple = ('全部指令', 0)  # 运行模式
        # 读取配置文件
        self.time_sleep = float(get_setting_data_from_ini('Config', '暂停时间'))
        self.branch_table_name: list = []
        # 互斥锁,用于暂停线程
        self.mutex = QMutex()
        self.condition = QWaitCondition()
        self.is_paused: bool = False

    def set_branch_name_index(self, branch_name_index_):
        """设置分支表名索引"""
        self.branch_name_index = branch_name_index_

    def set_run_mode(self, mode: str, info: int):
        """设置运行模式
        :param mode: 运行模式（全部指令、单行指令）
        :param info: 指令ID"""
        self.run_mode = (mode, info)

    def set_repeat_number(self, number: int):
        """设置循环次数
        :param number: 循环次数，-1为无限循环"""
        self.number_cycles = number

    def show_message(self, message):
        """显示消息"""
        self.send_message.emit(message)

    def run(self):
        """执行指令"""
        self.start_state = True
        self.suspended = False
        # 从数据库中获取要执行的指令列表，并设置不同的运行模式
        list_instructions: list = []
        current_index: int = 0
        # 检查 self.run_mode 是否为空
        if not self.run_mode:
            self.show_message("运行模式未设置")
            return
        # 不断尝试获取指令列表，直到成功
        while True:
            if self.run_mode[0] == '全部指令':
                list_instructions = extracted_ins_from_database()
                current_index = 0
            elif self.run_mode[0] == '单行指令':
                list_instructions = extracted_ins_target_id_from_database(self.run_mode[1])
                current_index = 0
            elif self.run_mode[0] == '从当前行运行':
                list_instructions = extracted_ins_from_database()
                current_index = self.run_mode[1]
            # 如果获取失败，等待一段时间再尝试
            if list_instructions is None:
                self.show_message("未能从数据库中获取指令，重试中...")
                time.sleep(0.1)  # 等待5秒再重试
            else:
                break
        # print('指令列表：', list_instructions)
        # 执行指令
        # 设置主流程循环前的参数
        loop_type = '无限循环' if self.number_cycles == -1 else '有限循环'
        self.number = 1
        # 开始循环执行指令
        while (self.start_state and loop_type == '无限循环') or \
                (loop_type == '有限循环' and self.number <= self.number_cycles):
            # 执行指令集中的指令
            self.execute_instructions(self.branch_name_index, current_index, list_instructions)
            self.show_message('换行')
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
            try:
                elem_ = list_instructions_[current_list_index][current_index]
                # 【指令集合【指令分支（指令元素[元素索引]）】】
                # print('执行当前指令：', elem_)
                # [
                #     [
                #         (13, None, '时间等待', "{'类型': '时间等待', '时长': 5, '单位': '秒'}", None, None, None, 1,
                #          '提示异常并暂停', '', '主流程'),
                #         (14, None, '坐标点击', "{'动作': '左键单击', '坐标': '1086-1414', '自定义次数': 0}", None, None,
                #          None, 1, '提示异常并暂停', '', '主流程')
                #     ],
                #     [
                #         (16, None, '鼠标点击',
                #          "{'鼠标': '左键', '次数': 1, '间隔': 100, '按压': 100}", None,
                #          None, None, 1, '提示异常并暂停', '', '分支1')
                #     ]
                # ]
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
                    # 命令类型与对应操作类的映射
                    command_mapping = {
                        "图像点击": (ImageClick, self.out_mes, dic_),
                        "多图点击": (MultipleImagesClick, self.out_mes, dic_),
                        "坐标点击": (CoordinateClick, self.out_mes, dic_),
                        "时间等待": (TimeWaiting, self.out_mes, dic_),
                        "图像等待": (ImageWaiting, self.out_mes, dic_),
                        "滚轮滑动": (RollerSlide, self.out_mes, dic_),
                        "文本输入": (TextInput, self.out_mes, dic_),
                        "移动鼠标": (MoveMouse, self.out_mes, dic_),
                        "按下键盘": (PressKeyboard, self.out_mes, dic_),
                        "中键激活": (MiddleActivation, self.out_mes, dic_),
                        "鼠标点击": (MouseClick, self.out_mes, dic_),
                        "信息录入": (InformationEntry, self.out_mes, dic_, self.number),
                        "打开网址": (OpenWeb, self.out_mes, dic_),
                        "元素控制": (EleControl, self.out_mes, dic_),
                        "网页录入": (WebEntry, self.out_mes, dic_, self.number),
                        "鼠标拖拽": (MouseDrag, self.out_mes, dic_),
                        "切换frame": (ToggleFrame, self.out_mes, dic_),
                        "保存表格": (SaveForm, self.out_mes, dic_),
                        "拖动元素": (DragWebElements, self.out_mes, dic_),
                        "屏幕截图": (FullScreenCapture, self.out_mes, dic_),
                        "切换窗口": (SwitchWindow, self.out_mes, dic_),
                        "发送消息": (SendWeChat, self.out_mes, dic_),
                        "数字验证码": (VerificationCode, self.out_mes, dic_),
                        "提示音": (PlayVoice, self.out_mes, dic_),
                        "倒计时窗口": (WaitWindow, self.out_mes, dic_),
                        "提示窗口": (DialogWindow, self.out_mes, dic_),
                        "跳转分支": (BranchJump, self.out_mes, dic_),
                        "终止流程": (TerminationProcess, self.out_mes, dic_),
                        "窗口控制": (WindowControl, self.out_mes, dic_),
                        "按键等待": (KeyWait, self.out_mes, dic_),
                        "获取时间": (GetTimeValue, self.out_mes, dic_),
                        "获取Excel": (GetExcelCellValue, self.out_mes, dic_, self.number),
                        "获取对话框": (GetDialogValue, self.out_mes, dic_),
                        "获取剪切板": (GetClipboard, self.out_mes, dic_),
                        "变量判断": (ContrastVariables, self.out_mes, dic_),
                        "运行Python": (RunPython, self.out_mes, dic_),
                        "运行cmd": (RunCmd, self.out_mes, dic_),
                        "运行外部文件": (RunExternalFile, self.out_mes, dic_),
                        "写入单元格": (InputCellExcel, self.out_mes, dic_, self.number),
                        "OCR识别": (TextRecognition, self.out_mes, dic_),
                        "获取鼠标位置": (GetMousePositon, self.out_mes, dic_),
                        "窗口焦点等待": (WindowFocusWait, self.out_mes, dic_),
                        "颜色判断": (ColorJudgment, self.out_mes, dic_),
                    }
                    # 根据命令类型执行相应操作
                    if cmd_type in command_mapping:
                        command_class, *args = command_mapping[cmd_type]
                        command_instance = command_class(*args)
                        self.show_message('换行')
                        self.show_message(f'执行ID为{str(dict(dic_)["ID"])}的指令：{cmd_type}')
                        command_instance.start_execute()

                    # 执行完毕后，跳转到下一条指令
                    current_index += 1

                except Exception as e:
                    info_e = str(e)
                    if not info_e:
                        info_e = str(type(e))
                    # except IndexError:
                    #     info_e = 'test'
                    str_id = str(dict(dic_)['ID'])

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
                            text=f'ID为{str_id}的指令执行异常！\n是否重试？\n\n错误类型：{info_e}',
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
                            text=f'ID为{str_id}的指令抛出异常！\n\n错误类型：{info_e}',
                            title='提示',
                            icon=pymsgbox.STOP
                        )
                        current_index += 1
                        self.start_state = False
                        break

                    # 终止所有任务
                    elif exception_handling == '终止所有任务':
                        system_prompt_tone('执行异常')
                        self.show_message(f'ID为{str_id}的指令触发‘终止流程’指令。')
                        current_index += 1
                        self.start_state = False
                        break

                    # 跳转分支指令
                    else:  # 跳转分支
                        self.show_message(f'转到分支：{exception_handling}')
                        target_branch_name = exception_handling.split('-')[0]  # 分支表名
                        # 目标分支表名在分支表名中的索引
                        self.branch_table_name = get_branch_info(True)
                        branch_table_name_index = self.branch_table_name.index(target_branch_name)
                        # 分支表中要跳转的指令索引
                        branch_ins_index = exception_handling.split('-')[1]
                        x = int(branch_table_name_index)
                        y = int(branch_ins_index) - 1
                        self.execute_instructions(x, y, list_instructions_)
                        break

            except IndexError:
                self.show_message(f'无法进行分支跳转')
                break
