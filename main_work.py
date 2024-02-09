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

#

from 数据库操作 import extract_global_parameter, extracted_ins_from_database


def exit_main_work():
    sys.exit()





class MainWork:
    """主要工作类"""
    sigkeyhot = pyqtSignal(str, name='sigkeyhot')  # 自定义信号，用于快捷键

    def __init__(self, main_window, navigation):
        # 终止和暂停标志
        self.start_state = True
        self.suspended = False
        self.main_window = main_window  # 主窗体
        self.navigation = navigation  # 导航窗体
        # 指令执行线程
        self.command_thread = CommandThread(self.main_window, self.navigation)
        self.command_thread.send_message.connect(self.send_message)
        self.command_thread.finished_signal.connect(self.thread_finished)
        # 信息窗口
        self.info = Info(self.main_window)

    def start_work(self, branch_name=None):
        """主要工作"""

        # def info_show():
        #     """显示信息窗口"""
        #     info = Info(self.main_window)  # 运行提示窗口
        #     info.show()
        #     QApplication.processEvents()
        #     return info

        # 开始执行主要操作
        self.info.show()
        QApplication.processEvents()  # 显示信息窗口
        self.command_thread.start()  # 启动执行线程




if __name__ == '__main__':
    pass
