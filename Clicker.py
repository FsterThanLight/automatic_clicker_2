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
from __future__ import print_function

import os.path
import shutil

import openpyxl
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtCore import QUrl, Qt
from PyQt5.QtGui import QDesktopServices, QPixmap, QFont
from PyQt5.QtWidgets import (QMainWindow, QTableWidgetItem, QHeaderView,
                             QDialog, QInputDialog, QMenu, QFileDialog, QStyle, QStatusBar, QMessageBox, QApplication,
                             QAction, QSplashScreen)
from openpyxl.utils import get_column_letter
from system_hotkey import SystemHotkey

from icon import Icon
from main_work import CommandThread
from 功能类 import close_browser
from 导航窗口功能 import Na
from 数据库操作 import *
from 窗体.about import Ui_About
from 窗体.mainwindow import Ui_MainWindow
from 设置窗口 import Setting
from 资源文件夹窗口 import Global_s
from 选择窗体 import Branch_exe_win

import collections
collections.Iterable = collections.abc.Iterable

# todo：RGB颜色检测功能
# todo: 验证码指令使用云码平台
# todo: 指令可编译为python代码
# todo: 从微信获取变量
# todo: 可暂时禁用指令功能
# todo: win通知指令
# todo: excel指令集
# todo: 调试模式
# todo: 动作录制功能
# todo: 使用将指定标题的窗口正常显示后会出现菜单栏阴影的问题
# todo: 输出信息用绿色和红色区分开时间和内容

# activate clicker

# pyinstaller -D -w -i clicker.ico Clicker.py --hidden-import=pyttsx4.drivers
# pyinstaller -D -i clicker.ico Clicker.py --hidden-import=pyttsx4.drivers

# 添加指令的步骤：
# 1. 在导航页的页面中添加指令的控件
# 2. 在导航页的页面中添加指令的处理函数
# 3. 在导航页的treeWidget中添加指令的名称
# 4. 在功能类中添加运行功能


OUR_WEBSITE = 'https://gitee.com/fasterthanlight/automatic_clicker/releases'
QQ = '308994839'
QQ_GROUP = 'https://qm.qq.com/q/3ih3PE16Mg'
CURRENT_VERSION = 'v0.25 Beta'


class Main_window(QMainWindow, Ui_MainWindow):
    """主窗口"""
    sigkeyhot = pyqtSignal(str, name='sigkeyhot')  # 自定义信号,用于快捷键

    def __init__(self):
        super().__init__()
        # 初始化窗体
        self.setupUi(self)
        # 窗口和信息
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)  # 实例化状态栏
        self.icon = Icon()  # 实例化图标
        self.add_recent_to_fileMenu()  # 将最近文件添加到菜单中
        self.branch_win = Branch_exe_win(self)  # 分支选择窗口
        # 显示导不同的窗口
        self.pushButton.clicked.connect(lambda: self.show_windows('导航'))  # 显示导航窗口
        self.pushButton_3.clicked.connect(lambda: self.show_windows('全局'))  # 显示全局参数窗口
        self.actions_2.triggered.connect(lambda: self.show_windows('设置'))  # 打开设置
        self.actionabout.triggered.connect(lambda: self.show_windows('关于'))  # 打开关于窗体
        self.actionhelp.triggered.connect(lambda: self.show_windows('说明'))  # 打开使用说明
        self.actionk.triggered.connect(lambda: self.show_windows('快捷键说明'))  # 打开快捷键说明
        # 主窗体表格功能
        self.actionx.triggered.connect(lambda: self.save_data('自动保存'))  # 保存指令数据
        self.actionb.triggered.connect(lambda: self.save_data('db'))  # 导出数据，导出按钮
        self.actiona.triggered.connect(lambda: self.save_data('excel'))  # 导出数据，导出按钮
        self.actionf.triggered.connect(lambda: self.data_import('资源文件夹路径'))  # 导入数据
        # 主窗体开始按钮
        self.pushButton_5.clicked.connect(self.start)
        self.start_time = None
        self.pushButton_4.clicked.connect(lambda: self.show_windows('分支选择'))  # 结束任务按钮
        self.pushButton_6.clicked.connect(lambda: self.global_shortcut_key('终止线程'))  # 结束任务按钮
        self.pushButton_7.clicked.connect(lambda: self.global_shortcut_key('暂停和恢复线程'))  # 暂停和恢复按钮
        self.toolButton_8.clicked.connect(self.exporting_operation_logs)  # 导出日志按钮
        self.load_branch_to_combobox()  # 加载分支列表
        # 创建和删除分支
        self.toolButton_2.clicked.connect(self.create_branch)
        self.toolButton.clicked.connect(self.delete_branch)
        self.comboBox.currentIndexChanged.connect(self.get_data)
        # 右键菜单
        self.tableWidget.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.generateMenu)
        # 指令执行线程
        self.command_thread = CommandThread(self, None)
        self.command_thread.send_message.connect(self.send_message)
        self.command_thread.finished_signal.connect(self.thread_finished)
        # 设置全局快捷键,用于执行指令的终止
        self.sigkeyhot.connect(self.global_shortcut_key)
        self.hk_stop = SystemHotkey()
        # 加载上次的指令表格
        self.get_data()
        self.tableWidget.installEventFilter(self)  # 安装事件过滤器,重新设置表格的快捷键
        # 加载窗体初始值
        self.load_initialization()

    def load_initialization(self):
        """加载窗体初始值"""
        set_window_size(self)  # 获取上次退出时的窗口大小
        # 显示工具栏
        judge = eval(get_setting_data_from_db('显示工具栏'))
        self.toolBar.setVisible(judge)
        self.actiong.setChecked(judge)
        # 注册全局快捷键
        try:
            self.hk_stop.register(('f11',), callback=lambda x: self.sendkeyevent("终止线程"))
            self.hk_stop.register(('f10',), callback=lambda x: self.sendkeyevent("开始线程"))
            self.hk_stop.register(('alt', 'f11',), callback=lambda x: self.sendkeyevent("暂停和恢复线程"))
            self.hk_stop.register(('alt', '1',), callback=lambda x: self.sendkeyevent("弹出分支选择窗口"))
        except Exception as e:
            print(e)
            QMessageBox.critical(self, "错误", "全局快捷键已失效！")
            sys.exit()

        # 设置状态栏信息
        self.statusBar.showMessage(f'软件版本：{CURRENT_VERSION}准备就绪...', 3000)

    def add_recent_to_fileMenu(self):
        """将最近文件添加到菜单中"""
        recently_opened_list = get_recently_opened_file('文件列表')
        current_file_path = get_setting_data_from_db('当前文件路径')
        # 将最近打开文件添加到菜单中
        if len(recently_opened_list) != 0:
            for file in recently_opened_list:
                file_action = QAction(text=file, parent=self)
                # 设置信号
                file_action.triggered.connect(
                    lambda checked, file_=file: self.open_recent_file(file_))
                file_action.setCheckable(True)
                # 设置当前文件为选中状态
                if file == current_file_path:
                    file_action.setChecked(True)
                self.menuzv.addAction(file_action)
        # 关闭菜单栏
        self.menuzv.close()

    def open_recent_file(self, file_path):
        """打开最近打开的文件
        :param file_path: 文件路径"""
        recent_file = get_setting_data_from_db('当前文件路径')
        if file_path != recent_file:
            if os.path.exists(file_path):
                self.data_import(file_path)
            elif not os.path.exists(file_path):
                # 如果文件不存在，则删除最近打开文件列表中的文件
                remove_recently_opened_file(file_path)
                # 从菜单中删除文件
                for action in self.menuzv.actions():
                    if action.text() == file_path:
                        self.menuzv.removeAction(action)
                QMessageBox.critical(self, "错误", "文件不存在！已经从最近打开文件中删除。")
        else:
            for action in self.menuzv.actions():
                if action.text() == file_path:
                    action.setChecked(True)

    def delete_data(self):
        """删除选中的数据行"""
        # 获取选中值的行号和id
        try:
            row = self.tableWidget.currentRow()
            xx = int(self.tableWidget.item(row, 7).text())
            # 删除数据库中指定id的数据
            cursor, con = sqlitedb()
            branch_name = self.comboBox.currentText()
            cursor.execute('delete from 命令 where ID=? and 隶属分支=?', (xx, branch_name,))
            con.commit()
            close_database(cursor, con)
            self.get_data(row)  # 调用get_data()函数，刷新表格
            # 状态栏显示信息
            self.statusBar.showMessage(f'删除指令。', 1000)
        except AttributeError:
            pass

    def copy_data(self):
        """复制指定id的指令数据，插入到对应的id位置"""

        def get_new_order():
            """获取新的指令数据"""
            branch_name = self.comboBox.currentText()
            cursor.execute('SELECT * FROM 命令 WHERE ID=? AND 隶属分支=?', (id_, branch_name,))
            list_order = cursor.fetchone()
            new_id_ = int(list_order[0]) + 1  # 获取id
            return (new_id_,) + list_order[1:10] + (branch_name,)

        try:
            row = self.tableWidget.currentRow()
            id_ = int(self.tableWidget.item(row, 7).text())  # 指令ID
            cursor, con = sqlitedb()
            new_list_order = get_new_order()  # 获取新的指令数据
            try:
                cursor.execute('INSERT INTO 命令 VALUES (?,?,?,?,?,?,?,?,?,?,?)', new_list_order)
                con.commit()
            except sqlite3.IntegrityError:
                # 如果下一个id已经存在，则将后面的id全部加1
                max_id_ = 1000000
                cursor.execute('UPDATE 命令 SET ID=ID+? WHERE ID>?', (max_id_, id_))
                cursor.execute('UPDATE 命令 SET ID=ID-? WHERE ID>?', (max_id_ - 1, max_id_ + int(id_)))
                cursor.execute('INSERT INTO 命令 VALUES (?,?,?,?,?,?,?,?,?,?,?)', new_list_order)
                con.commit()
            close_database(cursor, con)
            self.get_data(row)
            self.statusBar.showMessage(f'复制指令。', 1000)
        except AttributeError:
            pass

    def go_to_branch(self):
        """转到分支"""
        row = self.tableWidget.currentRow()  # 获取当前行行号
        branch_name = self.tableWidget.item(row, 2).text()  # 分支名称
        if branch_name not in ['自动跳过', '提示异常并暂停', '提示异常并停止']:
            # 跳转到对应的分支表和行
            go_branch_name = branch_name.split('-')[0]
            go_row_num = branch_name.split('-')[1]
            self.comboBox.setCurrentText(go_branch_name)
            self.tableWidget.setCurrentCell(int(go_row_num) - 1, 0)  # 设置焦点
        else:
            self.statusBar.showMessage(f'当前指令无分支。', 1000)

    def modify_parameters(self):
        """修改参数"""
        try:
            # 获取当前行行号列号
            row = self.tableWidget.currentRow()
            id_ = int(self.tableWidget.item(row, 7).text())  # 指令ID
            ins_type = self.tableWidget.item(row, 1).text()  # 指令类型
            # 将导航页的tabWidget设置为对应的页
            navigation = Na(self)  # 实例化导航页窗口
            # 修改数据中的参数
            navigation.pushButton_2.setText('修改指令')
            navigation.modify_id = id_
            navigation.show()
            navigation.switch_navigation_page(ins_type)
        except AttributeError:
            QMessageBox.information(self, "提示", "请先选择一行待修改的数据！")

    def generateMenu(self, pos):
        """生成右键菜单"""

        def clear_table():
            """清空表格和数据库"""
            choice = QMessageBox.question(self, "提示", "确认清除所有指令吗？")
            if choice == QMessageBox.Yes:
                clear_all_ins()
                self.get_data()
            else:
                pass

        def insert_data_before(judge):
            """在目标指令前插入指令
            :param judge: （向前插入、向后插入）"""
            try:
                # 获取当前行行号列号
                row = self.tableWidget.currentRow()
                target_id = int(self.tableWidget.item(row, 7).text())  # 指令ID
                navigation = Na(self)  # 实例化导航页窗口
                navigation.show()
                # 修改数据中的参数
                navigation.pushButton_2.setText(judge)
                navigation.modify_id = target_id
                navigation.modify_row = row
            except AttributeError:
                QMessageBox.information(self, "提示", "请先选择一行待修改的数据！")

        # 表格右键菜单
        row_num = -1
        for i_ in self.tableWidget.selectionModel().selection().indexes():
            row_num = i_.row()
        if row_num != -1:  # 未选中数据不弹出右键菜单
            menu = QMenu()  # 实例化菜单

            refresh = menu.addAction("刷新")
            refresh.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))  # 设置图标

            up_ins = menu.addAction("上移")
            up_ins.setShortcut('Shift+↑')
            up_ins.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp))  # 设置图标

            down_ins = menu.addAction("下移")
            down_ins.setShortcut('Shift+↓')
            down_ins.setIcon(self.style().standardIcon(QStyle.SP_ArrowDown))  # 设置图标

            menu.addSeparator()
            insert_ins_before = menu.addAction("在前面插入指令")
            insert_ins_before.setIcon(self.icon.move_up)  # 设置图标

            insert_ins_after = menu.addAction("在后面插入指令")
            insert_ins_after.setIcon(self.icon.move_down)  # 设置图标

            menu.addSeparator()
            copy_ins = menu.addAction("复制指令")
            copy_ins.setShortcut('Ctrl+C')
            copy_ins.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))  # 设置图标

            modify_ins = menu.addAction("修改指令")
            modify_ins.setShortcut('Ctrl+Y')
            modify_ins.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))  # 设置图标

            go_branch = menu.addAction("转到分支")
            go_branch.setShortcut('Ctrl+G')
            go_branch.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))  # 设置图标

            del_ins = menu.addAction("删除指令")
            del_ins.setShortcut('Delete')
            del_ins.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))  # 设置图标

            menu.addSeparator()
            del_branch = menu.addAction("删除当前分支指令")
            del_branch.setIcon(self.style().standardIcon(QStyle.SP_DialogDiscardButton))  # 设置图标

            del_all_ins = menu.addAction("删除全部指令")
            del_all_ins.setIcon(self.icon.delete)  # 设置图标

            action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        else:
            return

        # 各项操作
        if action == copy_ins:
            self.copy_data()  # 复制指令
        elif action == del_ins:
            self.delete_data()  # 删除指令
        elif action == up_ins:
            self.go_up_down('up')
        elif action == down_ins:
            self.go_up_down('down')
        elif action == modify_ins:
            self.modify_parameters()  # 修改指令
        elif action == insert_ins_before:
            insert_data_before('向前插入')
        elif action == insert_ins_after:
            insert_data_before('向后插入')
        elif action == refresh:
            self.get_data()
            self.statusBar.showMessage(f'刷新指令表格。', 1000)
        elif action == del_all_ins:
            clear_table()
            self.statusBar.showMessage(f'清空指令表格。', 1000)
        elif action == del_branch:
            clear_all_ins(branch_name=self.comboBox.currentText())
            self.get_data()
            self.statusBar.showMessage(f'清空当前分支全部指令。', 1000)
        elif action == go_branch:
            self.go_to_branch()

    def show_windows(self, judge):
        """打开窗体"""
        if judge == '设置':
            setting_win = Setting(self)  # 设置窗体
            setting_win.setModal(True)
            setting_win.exec_()
        elif judge == '全局':
            global_s = Global_s(self)  # 全局设置窗口
            global_s.setModal(True)
            global_s.exec_()
        elif judge == '导航':
            navigation = Na(self)  # 实例化导航页窗口
            navigation.show()
        elif judge == '关于':
            about = About(self)  # 设置关于窗体
            about.setModal(True)
            about.exec_()
        elif judge == '分支选择':  # 分支选择窗口
            if not self.branch_win.isVisible():
                self.branch_win.show()
            else:
                self.branch_win.close()
        elif judge == '说明':
            QDesktopServices.openUrl(QUrl(OUR_WEBSITE))
        elif judge == '快捷键说明':
            # 使用MessageBox显示快捷键说明
            QMessageBox.information(
                self, "快捷键说明",
                "Ctrl+Enter：\t添加指令\n"
                "Ctrl+C：\t\t复制指令\n"
                "Delete：\t\t删除指令\n"
                "Shift+↑：\t上移指令\n"
                "Shift+↓：\t下移指令\n"
                "Ctrl+↑：\t\t切换到上个分支\n"
                "Ctrl+↓：\t\t切换到下个分支\n"
            )

    def get_data(self, row=None):
        """从数据库获取数据并存入表格
        :param row: 设置焦点行号"""
        try:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(0)
            # 获取数据库数据
            cursor, con = sqlitedb()
            branch_name = self.comboBox.currentText()
            cursor.execute(
                'select 图像名称,指令类型,异常处理,备注,参数1,参数2,重复次数,ID from 命令 where 隶属分支=?',
                (branch_name,)
            )
            list_order = cursor.fetchall()
            close_database(cursor, con)
            # 在表格中写入数据
            for i_ in range(len(list_order)):
                self.tableWidget.insertRow(i_)
                for j in range(len(list_order[i_])):
                    self.tableWidget.setItem(i_, j, QTableWidgetItem(str(list_order[i_][j])))
            # 自适应列宽（排除第一列）
            header = self.tableWidget.horizontalHeader()
            for col in range(1, header.count()):
                header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
            # 设置焦点
            if row is not None:
                self.tableWidget.setCurrentCell(int(row), 0)

        except sqlite3.OperationalError:
            pass

    def go_up_down(self, judge):
        """向上或向下移动选中的行"""

        def database_exchanges_two_rows(id_1: int, id_2: int) -> None:
            """交换数据库中的两行数据
            :param id_1: 要交换的第一行的id
            :param id_2: 要交换的第二行的id"""
            cursor, con = sqlitedb()
            # 交换两行的id
            cursor.execute('update 命令 set ID=? where ID=?', (999999, id_1,))
            cursor.execute('update 命令 set ID=? where ID=?', (id_1, id_2,))
            cursor.execute('update 命令 set ID=? where ID=?', (id_2, 999999,))
            con.commit()
            # 刷新数据库
            close_database(cursor, con)

        try:
            # 获取选中值的行号和id
            row = self.tableWidget.currentRow()
            id_ = int(self.tableWidget.item(row, 7).text())
            if judge == 'up':
                if row != 0:
                    # 查询上一行的id
                    id_row_up = int(self.tableWidget.item(row - 1, 7).text())
                    # 交换两行的id
                    database_exchanges_two_rows(id_, id_row_up)
                    self.get_data(row - 1)
                    self.statusBar.showMessage(f'上移指令。', 1000)
            elif judge == 'down':
                if row != self.tableWidget.rowCount() - 1:
                    # 查询下一行的id
                    id_row_down = int(self.tableWidget.item(row + 1, 7).text())
                    # 交换两行的id
                    database_exchanges_two_rows(id_, id_row_down)
                    self.get_data(row + 1)
                    self.statusBar.showMessage(f'下移指令。', 1000)
        except AttributeError:
            pass

    def save_data(self, judge: str):
        """保存配置文件到当前文件夹下
        :param judge: 保存的文件类型（db、excel、自动保存）"""

        def get_directory_pyth_global(judge__):
            """获取全局资源文件夹路径,如果不存在最近打开的文件路径"""
            resource_folder_path = extract_global_parameter('资源文件夹路径')
            directory_folder_path = resource_folder_path[0] if \
                resource_folder_path else os.path.expanduser("~")
            # 默认文件名
            default_file_name = '指令数据.xlsx' if judge__ == 'excel' else '指令数据.db'
            return os.path.normpath(os.path.join(directory_folder_path, default_file_name))

        def get_file_and_folder(judge__) -> tuple:
            """获取文件名和文件夹路径"""
            directory_path = get_directory_pyth_global(judge__)  # 获取全局资源文件夹路径
            # 打开选择文件对话框
            file_path, _ = QFileDialog.getSaveFileName(
                parent=self,
                caption="保存文件",
                filter="(*.db)" if judge__ == 'db' else "(*.xlsx)",
                directory=directory_path
            )
            if file_path != '':  # 获取文件名称
                return (
                    os.path.normpath(os.path.split(file_path)[1]),
                    os.path.normpath(os.path.split(file_path)[0])
                )
            else:
                return None, None

        def get_file_and_folder_from_setting():
            """获取文件名和文件夹路径"""
            # 获取资源文件夹路径，如果存在则使用用户的主目录
            recently_opened = get_setting_data_from_db('当前文件路径')
            if (recently_opened != 'None') and (os.path.exists(recently_opened)):
                return (
                    os.path.normpath(os.path.split(recently_opened)[1]),
                    os.path.normpath(os.path.split(recently_opened)[0])
                )
            else:
                self.statusBar.showMessage(f'未找到最近导入的文件路径。已自动切换为另存为...', 3000)
                return get_file_and_folder('excel')

        # 判断是否为另存为,如果不是则自动判断文件类型
        try:
            print('保存数据', judge)
            judge_ = None
            if judge != '自动保存':
                file_name, folder_path = get_file_and_folder(judge)  # 获取文件名和文件夹路径
                judge_ = judge
            else:
                file_name, folder_path = get_file_and_folder_from_setting()  # 如果存在最近打开的文件路径
                if (file_name != '') and (file_name is not None):
                    judge_ = 'excel' if file_name.endswith('.xlsx') else 'db'
            # 开始保存数据
            if (file_name is not None) and (folder_path is not None):
                if judge_ == 'db':
                    # 连接数据库
                    cursor, con = sqlitedb()
                    # 获取数据库文件路径
                    db_file = con.execute('PRAGMA database_list').fetchall()[0][2]
                    close_database(cursor, con)
                    # 将数据库文件复制到指定文件夹下
                    save_path = os.path.normpath(os.path.join(folder_path, file_name))
                    shutil.copy(db_file, save_path)
                    # 提示保存成功
                    if judge != '自动保存':
                        QMessageBox.information(self, "提示", "指令数据保存成功！")
                    self.statusBar.showMessage(f'指令数据已保存至{save_path}。', 3000)
                elif judge_ == 'excel':
                    # 使用openpyxl模块创建Excel文件
                    wb = openpyxl.Workbook()
                    # 获取全局参数表中的资源文件夹路径
                    branch_table_list = extract_global_parameter('分支表名')
                    # 将sheet名设置为分支表名
                    for branch_name in branch_table_list:
                        wb.create_sheet(branch_name)  # 创建所有分支sheet
                        # 向分支sheet中写入数据
                        sheet = wb[branch_name]
                        # 设置表头
                        sheet['A1'] = 'ID'
                        sheet['B1'] = '图像名称'
                        sheet['C1'] = '指令类型'
                        sheet['D1'] = '参数1'
                        sheet['E1'] = '参数2'
                        sheet['F1'] = '参数3'
                        sheet['G1'] = '参数4'
                        sheet['H1'] = '重复次数'
                        sheet['I1'] = '异常处理'
                        sheet['J1'] = '备注'
                        sheet['K1'] = '隶属分支'
                        # 写入数据
                        branch_list_instructions = extracted_ins_from_database(branch_name)
                        for ins in range(len(branch_list_instructions)):
                            for i_ in range(len(branch_list_instructions[ins])):
                                sheet.cell(row=ins + 2, column=i_ + 1, value=branch_list_instructions[ins][i_])
                        # 自适应列宽
                        for col in range(1, sheet.max_column + 1):
                            max_length = 0
                            for cell in sheet[get_column_letter(col)]:
                                cell_length = 0.7 * len(re.findall('([\u4e00-\u9fa5])',
                                                                   str(cell.value))) + len(str(cell.value))
                                max_length = max(max_length, cell_length)
                            sheet.column_dimensions[get_column_letter(col)].width = max_length + 5

                    wb.remove(wb['Sheet'])  # 删除默认的sheet
                    # 保存Excel文件
                    save_path = os.path.normpath(os.path.join(folder_path, file_name))
                    wb.save(save_path)
                    # 提示保存成功，是否打开文件夹
                    if judge != '自动保存':
                        choice__ = QMessageBox.question(self,
                                                        '提示',
                                                        '指令数据保存成功！是否打开文件夹？',
                                                        QMessageBox.Yes | QMessageBox.No,
                                                        QMessageBox.No
                                                        )
                        if choice__ == QMessageBox.Yes:
                            os.startfile(save_path)
                    self.statusBar.showMessage(f'指令数据已保存至{save_path}。', 3000)
        except PermissionError:
            QMessageBox.critical(self, "错误", "保存失败，文件被占用！")

    def closeEvent(self, event):
        """关闭窗口事件"""
        # 是否隐藏工具栏
        update_settings_in_database(显示工具栏=str(self.actiong.isChecked()))
        # 终止线程
        if self.command_thread.isRunning():
            self.command_thread.terminate()
        # 是否退出清空数据库
        if eval(get_setting_data_from_db('退出提醒清空指令')):
            choice = QMessageBox.question(self, "提示", "确定退出并清空所有指令？\n将自动保存当前指令数据。")
            if choice == QMessageBox.Yes:
                # 退出终止后台进程并清空数据库
                self.save_data('自动保存')
                event.accept()
                clear_all_ins()
            else:
                event.ignore()
        self.branch_win.close()  # 关闭选择窗口
        # 窗口大小
        save_window_size((self.width(), self.height()), self.windowTitle())

    def data_import(self, file_path):
        """导入数据功能"""

        def data_import_from_db(target_path_):
            # 获取当前文件夹路径
            # 将目标数据库中的数据导入到当前数据库中
            cursor, con = sqlitedb()
            # 获取目标数据库中的数据
            con_target = sqlite3.connect(target_path_[0])
            cursor_target = con_target.cursor()
            cursor_target.execute('select * from 命令')
            list_instructions = cursor_target.fetchall()
            # 获取目标数据库中的分支表名
            cursor_target.execute(f"select 分支表名 from 全局参数")
            branch_result_list = [item[0] for item in cursor_target.fetchall() if item[0] is not None]
            close_database(cursor_target, con_target)
            # 将数据导入到当前数据库中
            try:
                # 更新命令表
                for ins in list_instructions:
                    cursor.execute('insert into 命令 values (?,?,?,?,?,?,?,?,?,?,?)', ins)
                    con.commit()
                # 更新分支表
                for branch_name in branch_result_list:
                    global_write_to_database('分支表名', branch_name)
                self.load_branch_to_combobox()  # 重新加载分支列表
                if file_path == '资源文件夹路径':
                    QMessageBox.information(self, "提示", "指令数据导入成功！")
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "导入失败", "ID重复或格式错误！")
            close_database(cursor, con)

        def data_import_from_excel(target_path_):
            # 读取数据
            wb = openpyxl.load_workbook(target_path_[0])
            sheets = wb.worksheets  # 获取所有的sheet
            cursor_, con_ = sqlitedb()
            for sheet in sheets:
                global_write_to_database('分支表名', sheet.title)  # 添加分支表名
                max_row = sheet.max_row
                max_column = sheet.max_column
                # 向数据库中写入数据
                try:
                    for row in range(2, max_row + 1):
                        # 获取第一列数据
                        instructions = []
                        for column in range(1, max_column + 1):
                            # 获取单元格数据
                            data = sheet.cell(row, column).value
                            instructions.append(data)
                        cursor_.execute(
                            "INSERT INTO 命令(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,"
                            "重复次数,异常处理,备注,隶属分支) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                            instructions[0:11])
                        con_.commit()
                except Exception as e:
                    # 捕获并处理异常
                    QMessageBox.warning(self, f"导入失败", f"ID重复或格式错误！{e}")
            close_database(cursor_, con_)
            self.load_branch_to_combobox()  # 重新加载分支列表
            if file_path == '资源文件夹路径':
                QMessageBox.information(self, "提示", "指令数据导入成功！")

        # 获取资源文件夹路径，如果不存在则使用用户的主目录
        if file_path == '资源文件夹路径':
            resource_folder_path = extract_global_parameter('资源文件夹路径')[0]
            # 获取当前文件夹路径
            directory_path = resource_folder_path if \
                resource_folder_path else os.path.expanduser("~")
            # 打开选择文件对话框
            target_path = QFileDialog.getOpenFileName(
                parent=self,
                caption="请选择指令备份文件",
                directory=directory_path,
                filter="(*.db *.xlsx)"
            )
            # 判断是否选择了文件
            if target_path[0] != '':
                suffix = os.path.splitext(target_path[0])[1]  # 获取文件后缀
            else:
                return
        else:
            target_path = (file_path, '')
            suffix = os.path.splitext(file_path)[1]

        # 如果为.db文件
        if suffix == '.db':
            clear_all_ins(True)  # 清空原有数据
            data_import_from_db(target_path)
        # 如果为.xlsx文件
        elif suffix == '.xlsx':
            clear_all_ins(True)  # 清空原有数据，包括分支表
            data_import_from_excel(target_path)
        # 将最近导入的文件路径写入数据库,用于保存时自动设置路径
        update_settings_in_database(当前文件路径=os.path.normpath(target_path[0]))  # 写入当前文件路径
        writes_to_recently_opened_files(os.path.normpath(target_path[0]))  # 写入最近打开的文件
        self.statusBar.showMessage(f'指令数据导入成功！已自动设置保存路径。', 1000)
        self.menuzv.clear()  # 清空最近打开文件菜单
        self.add_recent_to_fileMenu()  # 将最近文件添加到菜单中

    def start(self):
        """主窗体开始按钮"""

        def operation_before_execution():
            """执行前的操作"""
            self.plainTextEdit.clear()  # 清空日志
            self.tabWidget.setCurrentIndex(0)  # 切换到日志页
            if self.checkBox_2.isChecked():
                self.hide()

        if not self.command_thread.isRunning():  # 如果线程未运行,则执行
            operation_before_execution()
            self.command_thread.set_branch_name_index(int(self.comboBox.currentIndex()))
            self.command_thread.start()
            # 记录开始时间的时间戳
            self.start_time = time.time()

    def exporting_operation_logs(self):
        """导出操作日志"""
        # 打开保存文件对话框
        target_path = QFileDialog.getSaveFileName(
            parent=self,
            caption="请选择保存路径",
            directory=os.path.join(os.path.expanduser("~"), "操作日志.txt"),
            filter="(*.txt)"
        )
        # 判断是否选择了文件
        if target_path[0] != '':
            # 获取操作日志
            logs = self.plainTextEdit.toPlainText()
            # 将操作日志写入文件
            with open(target_path[0], 'w') as f:
                f.write(f'日志导出时间：{get_str_now_time()}\n')
                f.write(logs)
            QMessageBox.information(self, "提示", "操作日志导出成功！")

    def create_branch(self):
        """创建分支表并重命名"""
        # 弹出输入对话框，提示输入分支名称
        text, ok = QInputDialog.getText(self, "创建分支", "请输入分支名称：")
        if ok:
            try:
                # 连接数据库
                cursor, con = sqlitedb()
                # 查找是否有同名分支
                cursor.execute('select 分支表名 from 全局参数 where 分支表名=?', (text,))
                x = cursor.fetchall()
                if len(x) > 0:
                    QMessageBox.information(self, "提示", "分支已存在！")
                    return
                else:
                    # 向全局参数表中添加分支表名
                    cursor.execute('insert into 全局参数(资源文件夹路径,分支表名) values(?,?)', (None, text))
                    con.commit()
                    # 弹出提示框，提示创建成功
                    QMessageBox.information(self, "提示", "分支创建成功！")
                # 关闭数据库连接
                close_database(cursor, con)
                # 加载分支
                self.load_branch_to_combobox()
            except sqlite3.OperationalError:
                QMessageBox.critical(self, "提示", "分支创建失败！")
                pass

    def delete_branch(self):
        """删除分支"""
        # 弹出输入对话框，提示输入分支名称
        print('删除分支')
        text = self.comboBox.currentText()
        if text == MAIN_FLOW:
            QMessageBox.information(self, "提示", "无法删除主分支！")
        else:
            # 将combox显示的名称切换为命令
            self.comboBox.setCurrentIndex(0)
            cursor, con = sqlitedb()
            cursor.execute('delete from 全局参数 where 分支表名=?', (text,))  # 删除分支名称
            con.commit()
            close_database(cursor, con)  # 关闭数据库连接
            QMessageBox.information(self, "提示", "分支删除成功！")
            self.load_branch_to_combobox()  # 重新加载分支列表

    def load_branch_to_combobox(self):
        """加载分支"""
        self.comboBox.clear()
        self.comboBox.addItems(extract_global_parameter('分支表名'))

    def eventFilter(self, obj, event):
        # 重写self.tableWidget的快捷键事件
        if obj == self.tableWidget:
            if event.type() == 6:  # 键盘按下事件
                # 如果按下delete键
                if event.key() == Qt.Key_Delete:
                    self.delete_data()
                # 如果按下ctrl+c键
                if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_C:
                    self.copy_data()
                # 如果按下shift+向上键
                if event.modifiers() == Qt.ShiftModifier and event.key() == Qt.Key_Up:
                    self.go_up_down('up')
                    # 将焦点下移一行,抵消上移的误差
                    self.tableWidget.setCurrentCell(self.tableWidget.currentRow() + 1, 0)
                # 如果按下shift+向下键
                if event.modifiers() == Qt.ShiftModifier and event.key() == Qt.Key_Down:
                    self.go_up_down('down')
                    # 将焦点上移一行,抵消下移的误差
                    self.tableWidget.setCurrentCell(self.tableWidget.currentRow() - 1, 0)
                # 如果按下ctrl+向上键
                if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Up:
                    # 将焦点下移一行,抵消上移的误差
                    self.tableWidget.setCurrentCell(self.tableWidget.currentRow() + 1, 0)
                    # 如果分支不为self.comboBox的第一个则,切换上一个分支
                    if self.comboBox.currentIndex() != 0:
                        self.comboBox.setCurrentIndex(self.comboBox.currentIndex() - 1)
                # 如果按下ctrl+向下键
                if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Down:
                    # 将焦点上移一行,抵消下移的误差
                    self.tableWidget.setCurrentCell(self.tableWidget.currentRow() - 1, 0)
                    # 如果分支不为self.comboBox的最后一个则,切换下一个分支
                    if self.comboBox.currentIndex() != self.comboBox.count() - 1:
                        self.comboBox.setCurrentIndex(self.comboBox.currentIndex() + 1)
                # 如果按下ctrl+g键
                if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_G:
                    self.go_to_branch()  # 转到分支
                # 如果按下ctrl+x键
                if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Y:
                    self.modify_parameters()  # 修改指令
        return super().eventFilter(obj, event)

        # 热键处理函数

    def global_shortcut_key(self, i_str):
        """全局热键处理函数"""
        system_prompt_tone('全局快捷键')  # 发出提示音

        if i_str == "终止线程":
            if self.command_thread.isRunning():
                self.command_thread.terminate()  # 终止线程
                # 获取当前时间
                self.plainTextEdit.appendPlainText(f'{get_str_now_time()} 任务终止！')
                if self.checkBox_2.isChecked():
                    self.show()
                QApplication.processEvents()
                show_normal_window_with_specified_title(self.windowTitle())  # 显示窗口

        elif i_str == "开始线程":
            self.start()  # 执行主任务
            self.plainTextEdit.appendPlainText(f'{get_str_now_time()} 任务开始！')

        elif i_str == "暂停和恢复线程":
            if self.command_thread.isRunning():
                if self.command_thread.is_paused:
                    self.plainTextEdit.appendPlainText(f'{get_str_now_time()} 任务恢复！')
                    self.command_thread.resume()
                else:
                    self.plainTextEdit.appendPlainText(f'{get_str_now_time()} 任务暂停！')
                    self.command_thread.pause()

        elif i_str == "弹出分支选择窗口":
            self.show_windows('分支选择')

    def sendkeyevent(self, i_str):
        """发送热键信号,将外部信号，转化成qt信号,用于全局热键"""
        self.sigkeyhot.emit(i_str)

    def send_message(self, message):
        """向日志窗口发送信息"""
        self.plainTextEdit.appendPlainText(f'{get_str_now_time()}\t{message}')

    def thread_finished(self, message):

        def send_elapsed_time():
            """发送耗时"""
            elapsed_time = time.time() - self.start_time
            # 将秒转换为毫秒或者保留两位小数的秒数
            if elapsed_time < 1:
                elapsed_time_ms = round(elapsed_time * 1000)  # 毫秒
                return f"{elapsed_time_ms}毫秒"
            else:
                elapsed_time_sec = round(elapsed_time, 2)  # 秒，保留两位小数
                return f"{elapsed_time_sec}秒"

        self.plainTextEdit.appendPlainText(f'{get_str_now_time()}\t{message}，耗时{send_elapsed_time()}。')
        if self.checkBox_2.isChecked():  # 显示窗口
            self.show()
            QApplication.processEvents()
        system_prompt_tone('线程结束')  # 发出提示音
        show_normal_window_with_specified_title(self.windowTitle())  # 显示窗口
        close_browser()  # 关闭浏览器驱动


class About(QDialog, Ui_About):
    """关于窗体"""

    def __init__(self, parent=None):
        super().__init__(parent)
        # 初始化窗体
        self.setupUi(self)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)  # 隐藏帮助按钮
        set_window_size(self)  # 获取上次退出时的窗口大小
        self.label_2.setText(f'版本：{CURRENT_VERSION}')  # 设置版本号
        self.label_7.setText('<a href="{}">{}</a>'.format(QQ_GROUP, QQ))
        # 绑定事件
        self.gitee.clicked.connect(self.show_gitee)

    @staticmethod
    def show_gitee():
        QDesktopServices.openUrl(QUrl(OUR_WEBSITE))

    def closeEvent(self, event):
        # 保存窗体大小
        save_window_size((self.width(), self.height()), self.windowTitle())


class QSSLoader:
    """QSS皮肤加载器"""

    def __init__(self):
        pass

    @staticmethod
    def read_qss_file(qss_file_name):
        """从文件中读取qss的静态方法"""
        with open(qss_file_name, "r", encoding="UTF-8") as file:
            return file.read()


if __name__ == "__main__":
    # 自适应高分辨率
    # QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    # app = QApplication([])
    app = QtWidgets.QApplication(sys.argv)

    splash = QSplashScreen(QPixmap(r'./flat/开屏.png'))  # 创建启动界面
    splash.showMessage(  # 初始文本
        "加载中......", QtCore.Qt.AlignHCenter | QtCore.Qt.AlignBottom, QtCore.Qt.green
    )
    splash.setFont(QFont('微软雅黑', 15))  # 设置字体
    splash.show()  # 显示启动界面
    time.sleep(0.25)  # 延时1秒

    main_win = Main_window()  # 创建主窗体

    # 设置窗体样式
    try:
        style_name = 'Combinear'
        style_file = r"./flat/{}.qss".format(style_name)
        style_sheet = QSSLoader.read_qss_file(style_file)
        main_win.setStyleSheet(style_sheet)
    except FileNotFoundError:
        pass

    main_win.show()  # 显示窗体，并根据设置检查更新

    splash.finish(main_win)  # 隐藏启动界面
    splash.deleteLater()

    sys.exit(app.exec_())

    # def is_admin():
    #     try:
    #         return ctypes.windll.shell32.IsUserAnAdmin()
    #     except:
    #         return False
    #
    # if is_admin():
    #     app = QApplication([])
    #     # 创建主窗体
    #     main_window = Main_window()
    #     # 显示窗体，并根据设置检查更新
    #     main_window.main_show()
    #     # 显示添加对话框窗口
    #     sys.exit(app.exec_())
    # else:
    #     if sys.version_info[0] == 3:
    #         ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    #     else:  # in python2.x
    #         ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(__file__), None, 1)
