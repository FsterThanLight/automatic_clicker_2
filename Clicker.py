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

import ctypes
import datetime
import json
import os
import re
import shutil
import sqlite3
import sys
import time
import webbrowser

import cryptocode
import keyboard
import openpyxl
import pyautogui
import requests
from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices, QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, \
    QFileDialog, QTableWidgetItem, QMessageBox, QHeaderView, QDialog, QInputDialog
from openpyxl.utils.exceptions import InvalidFileException
from pyscreeze import unicode

from main_work import MainWork, exit_main_work
from 窗体.about import Ui_Dialog
from 窗体.global_s import Ui_Global
from 窗体.info import Ui_Form
from 窗体.mainwindow import Ui_MainWindow
from 窗体.navigation import Ui_navigation
from 窗体.setting import Ui_Setting
# 截图模块
from screen_capture import ScreenCapture
# 网页自动化模块
import selenium

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'}


def load_json():
    """从json文件中加载更新网址和保留文件名"""
    file_name = 'update_data.json'
    with open(file_name, 'r', encoding='utf8') as f:
        data = json.load(f)
    url = cryptocode.decrypt(data['url_encrypt'], '123456')
    # print(url)
    return url


def get_download_address(main_window, warning):
    """获取下载地址、版本信息、更新说明"""
    global headers
    url = load_json()
    # print(url)
    try:
        res = requests.get(url, headers=headers, timeout=0.2)
        info = cryptocode.decrypt(res.text, '123456')
        list_1 = info.split('=')
        # print(list_1)
        return list_1
    except requests.exceptions.ConnectionError:
        if warning == 1:
            # print("无法获取更新信息，请检查网络。")
            QMessageBox.critical(main_window, "更新检查", "无法获取更新信息，请检查网络。")
            time.sleep(1)
        else:
            pass


class Main_window(QMainWindow, Ui_MainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        # 初始化窗体
        self.setupUi(self)
        # 软件版本
        self.version = 'v0.21'
        # 窗体的功能
        self.main_work = MainWork(self)
        # 全局设置窗口
        self.global_s = Global_s()
        # 实例化导航页窗口
        self.navigation = Na(self.global_s)
        # 实例化设置窗口
        self.setting = Setting()
        # 设置关于窗体
        self.about = About()
        # 提示窗口
        self.info = Info()
        # 设置表格列宽自动变化，并使第5列列宽固定
        self.format_table()
        # 显示导航页窗口
        self.pushButton.clicked.connect(self.show_navigation)
        # 显示全局参数窗口
        self.pushButton_3.clicked.connect(self.show_global_s)
        # 获取数据，修改按钮
        self.toolButton_5.clicked.connect(self.get_data)
        # 获取数据，子窗体取消按钮
        # self.dialog_1.pushButton_2.clicked.connect(self.get_data)
        self.navigation.pushButton_3.clicked.connect(self.get_data)
        # 获取数据，子窗体保存按钮
        # self.dialog_1.pushButton.clicked.connect(self.get_data)
        self.navigation.pushButton_2.clicked.connect(self.get_data)
        # 删除数据，删除按钮
        self.pushButton_2.clicked.connect(self.delete_data)
        # 交换数据，上移按钮
        self.toolButton_3.clicked.connect(lambda: self.go_up_down("up"))
        self.toolButton_4.clicked.connect(lambda: self.go_up_down("down"))
        # 导出数据，导出按钮
        self.actionb.triggered.connect(self.save_data_to_current)
        # 清空指令按钮
        self.toolButton_6.clicked.connect(self.clear_table)
        # 导入数据按钮
        self.actionf.triggered.connect(self.data_import)
        # 主窗体开始按钮
        self.pushButton_5.clicked.connect(self.start)
        self.pushButton_4.clicked.connect(lambda: self.start(only_current_instructions=True))
        # 打开设置
        self.actions_2.triggered.connect(self.show_setting)
        # 结束任务按钮
        self.pushButton_6.clicked.connect(exit_main_work)
        # 导出日志按钮
        self.toolButton_8.clicked.connect(self.exporting_operation_logs)
        # 检查更新按钮（菜单栏）
        self.actionj.triggered.connect(lambda: self.check_update(1))
        # 隐藏工具栏
        self.actiong.triggered.connect(self.hide_toolbar)
        # 打开关于窗体
        self.actionabout.triggered.connect(self.show_about)
        # 打开使用说明
        self.actionhelp.triggered.connect(self.open_readme)
        # 修改指令按钮
        self.tab_index = {
            "图像点击": 0,
            "坐标点击": 1,
            "鼠标移动": 2,
            "等待": 3,
            "滚轮滑动": 4,
            "文本输入": 5,
            "按下键盘": 6,
            "中键激活": 7,
            "鼠标事件": 8,
            "excel信息录入": 9
        }
        self.pushButton_8.clicked.connect(self.modify_parameters)

        # 分支表名
        self.branch_name = []
        self.load_branch()
        # 创建和删除分支
        self.toolButton_2.clicked.connect(self.create_branch)
        self.toolButton.clicked.connect(self.delete_branch)
        self.comboBox.currentIndexChanged.connect(self.get_data)

    def sqlitedb(self):
        """建立与数据库的连接，返回游标"""
        try:
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            return cursor, con
        except sqlite3.Error:
            print("数据库连接失败")
            sys.exit()

    def close_database(self, cursor, conn):
        """关闭数据库"""
        cursor.close()
        conn.close()

    def format_table(self):
        """设置主窗口表格格式"""
        # 列的大小拉伸，可被调整
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 列的大小为可交互式的，用户可以调整
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.tableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.tableWidget.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.tableWidget.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        # 列的大小调整为固定，列宽不会改变
        self.tableWidget.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
        self.tableWidget.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed)
        # 设置列宽为50像素
        self.tableWidget.setColumnWidth(6, 30)
        self.tableWidget.setColumnWidth(7, 30)

    def show_setting(self):
        self.setting.show()
        self.setting.load_setting_data()
        print('设置窗口打开')
        resize = self.geometry()
        self.setting.move(resize.x() + 90, resize.y())

    def show_about(self):
        """显示关于窗口"""
        self.about.show()
        print('关于窗体开启')
        resize = self.geometry()
        self.about.move(resize.x() + 90, resize.y())

    def show_navigation(self):
        self.navigation.show()
        # 加载导航页数据
        self.navigation.load_values_to_controls()
        print("导航页窗口开启")
        resize = self.geometry()
        self.setting.move(resize.x() + 90, resize.y())

    def show_global_s(self):
        self.global_s.show()
        print("全局参数窗口开启")
        resize = self.geometry()
        self.setting.move(resize.x() + 90, resize.y())

    def get_data(self):
        """从数据库获取数据并存入表格"""
        print('刷新表格')
        try:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(0)
            # 获取数据库数据
            cursor, con = self.sqlitedb()
            branch_name = self.comboBox.currentText()
            cursor.execute(
                'select 图像名称,指令类型,异常处理,备注,参数1,参数2,重复次数,ID from 命令 where 隶属分支=?',
                (branch_name,))
            # cursor.execute('select 图像名称,指令类型,异常处理,参数1,参数2,参数3,参数4,重复次数,ID from 命令')
            list_order = cursor.fetchall()
            self.close_database(cursor, con)
            # 在表格中写入数据
            for i in range(len(list_order)):
                self.tableWidget.insertRow(i)
                for j in range(len(list_order[i])):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(list_order[i][j])))
        except sqlite3.OperationalError:
            pass

    def delete_data(self):
        """删除选中的数据行"""
        # 获取选中值的行号和id
        try:
            row = self.tableWidget.currentRow()
            column = self.tableWidget.currentColumn()
            xx = self.tableWidget.item(row, 7).text()
            print(row, column, xx)
            # 删除数据库中指定id的数据
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            branch_name = self.comboBox.currentText()
            cursor.execute('delete from 命令 where ID=? and 隶属分支=?', (xx, branch_name,))
            con.commit()
            con.close()
            # 调用get_data()函数，刷新表格
            self.get_data()
        except AttributeError:
            pass

    def go_up_down(self, judge):
        """向上或向下移动选中的行"""
        # 获取选中值的行号和id
        row = self.tableWidget.currentRow()
        column = self.tableWidget.currentColumn()
        try:
            xx = self.tableWidget.item(row, 7).text()
            # 将选中行的数据在数据库中与上一行数据交换，如果是第一行则不交换
            id = int(self.tableWidget.item(row, 7).text())
            # 初始化值
            id_up_down = id
            row_up_down = row
            # 判断是否执行数据库操作
            execute_sql = False
            # 判断是向上还是向下移动
            if judge == 'up':
                if row != 0:
                    # 获取选中值的行号
                    id_up_down = id - 1
                    row_up_down = row - 1
                    execute_sql = True
            elif judge == 'down':
                if row != self.tableWidget.rowCount() - 1:
                    # 获取选中值的行号
                    id_up_down = id + 1
                    row_up_down = row + 1
                    execute_sql = True
            if execute_sql:
                # 连接数据库
                print("执行数据库操作")
                con = sqlite3.connect('命令集.db')
                cursor = con.cursor()
                # 获取选中行和上一行的数据
                branch_name = self.comboBox.currentText()
                cursor.execute(
                    'select 图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支,ID from 命令 where ID=? and 隶属分支=?',
                    (id, branch_name,))
                list_id = cursor.fetchall()
                cursor.execute(
                    'select 图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支,ID from 命令 where ID=? and 隶属分支=?',
                    (id_up_down, branch_name,))
                list_id_up = cursor.fetchall()
                # 交换选中行和上一行的数据
                cursor.execute(
                    'update 命令 set 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=?,备注=?,隶属分支=? where ID=? and '
                    '隶属分支=?',
                    (list_id_up[0][0], list_id_up[0][1], list_id_up[0][2], list_id_up[0][3], list_id_up[0][4],
                     list_id_up[0][5], list_id_up[0][6], list_id_up[0][7], list_id_up[0][8], list_id_up[0][9], id,
                     branch_name,))
                cursor.execute(
                    'update 命令 set 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=?,备注=?,隶属分支=? where ID=? and '
                    '隶属分支=?',
                    (list_id[0][0], list_id[0][1], list_id[0][2], list_id[0][3], list_id[0][4], list_id[0][5],
                     list_id[0][6], list_id[0][7], list_id[0][8], list_id[0][9], id_up_down, branch_name,))
                con.commit()
                con.close()
            # 调用get_data()函数，刷新表格
            self.get_data()
            # 将焦点移动到交换后的行
            self.tableWidget.setCurrentCell(row_up_down, column)
        except AttributeError:
            pass

    def save_data_to_current(self):
        """保存配置文件到当前文件夹下"""
        # 打开选择文件夹对话框
        target_path = QFileDialog.getExistingDirectory(self, "选择保存路径。")
        # 弹出输入框，获取文件名
        file_name, ok = QInputDialog.getText(self, "保存文件", "请输入保存指令的文件名：")
        if ok:
            # 连接数据库
            con = sqlite3.connect('命令集.db')
            # 获取数据库文件路径
            db_file = con.execute('PRAGMA database_list').fetchall()[0][2]
            con.close()
            # 判断是否输入文件名
            if file_name == '':
                QMessageBox.warning(self, "警告", "请输入文件名！")
            else:
                # 判断是否选择了文件夹
                if target_path == '':
                    QMessageBox.warning(self, "警告", "请选择保存路径！")
                else:
                    # 将数据库文件复制到指定文件夹下
                    shutil.copy(db_file, target_path + '/' + file_name + '.db')
                    QMessageBox.information(self, "提示", "指令数据保存成功！")

    def clear_database(self):
        """清空数据库"""
        cursor, con = self.sqlitedb()
        # 清空分支列表中所有的数据
        cursor.execute('delete from 命令 where ID<>-1')
        con.commit()
        self.close_database(cursor, con)

    def closeEvent(self, event):
        choice = QMessageBox.question(self, "提示", "确定退出并清空所有指令？")
        if choice == QMessageBox.Yes:
            # 退出终止后台进程并清空数据库
            event.accept()
            self.clear_database()
            exit_main_work()
        else:
            event.ignore()

    def clear_table(self):
        """清空表格和数据库"""
        choice = QMessageBox.question(self, "提示", "确认清除所有指令吗？")
        if choice == QMessageBox.Yes:
            self.clear_database()
            self.get_data()
        else:
            pass

    def data_import(self):
        """导入数据功能"""
        # 打开选择文件对话框
        target_path = QFileDialog.getOpenFileName(self, "请选择指令备份文件", '', "(*.db)")
        # 判断是否选择了文件
        if target_path[0] == '':
            pass
        else:
            # 获取当前文件夹路径
            cwd = os.getcwd()
            # 复制数据库文件到当前文件夹下，并将其重命名为'命令集.db'取代原有数据库文件
            shutil.copy(target_path[0], cwd + '/命令集.db')
            QMessageBox.information(self, "提示", "指令数据导入成功！")
            self.load_branch()

    def start(self, only_current_instructions=False):
        """主窗体开始按钮"""

        def info_show():
            """显示信息窗口"""
            self.info.show()
            resize = self.geometry()
            self.info.move(resize.x() + 45, resize.y() - 30)

        # 开始主任务
        if not only_current_instructions:
            info_show()
            self.main_work.start_work()
        elif only_current_instructions:
            if self.comboBox.currentText() == '主流程':
                QMessageBox.warning(self, "警告", "主分支无法执行该操作！")
            else:
                info_show()
                self.main_work.start_work(only_current_instructions)
        self.info.close()

    def clear_plaintext(self, judge):
        """清空处理框中的信息"""
        if judge == 200:
            lines = self.plainTextEdit.blockCount()
            if lines > 200:
                self.plainTextEdit.clear()
        else:
            self.plainTextEdit.clear()

    def check_update(self, warning):
        """检查更新功能"""
        pass
        # 获取下载地址、版本号、更新信息
        list_1 = get_download_address(self, warning)
        # print(list_1)
        try:
            address = list_1[0]
            version = list_1[1]
            information = list_1[2]
            # 判断是否有更新
            print(version)
            if version != self.version:
                x = QMessageBox.information(self, "更新检查",
                                            "已发现最新版" + version + "\n是否手动下载最新安装包？" + '\n' + information,
                                            QMessageBox.Yes | QMessageBox.No,
                                            QMessageBox.Yes)
                if x == QMessageBox.Yes:
                    # 打开下载地址
                    webbrowser.open(address)
                    # os.popen('update.exe')
                    sys.exit()
            else:
                if warning == 1:
                    QMessageBox.information(self, "更新检查", "当前" + self.version + "已是最新版本。")
                else:
                    pass
        except TypeError:
            pass

    def main_show(self):
        """显示窗体，并根据设置检查更新"""
        self.show()
        # import sqlite3
        # 连接数据库获取是否检查更新选项
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select 值 from 设置 where 设置类型=?', ('启动检查更新',))
        x = cursor.fetchall()[0][0]
        cursor.close()
        print('启动检查更新')
        print(x)
        if x == 1:
            self.check_update(0)
        else:
            pass

    def hide_toolbar(self):
        """隐藏工具栏"""
        if self.actiong.isChecked():
            self.toolBar.show()
        elif not self.actiong.isChecked():
            self.toolBar.hide()

    def exporting_operation_logs(self):
        """导出操作日志"""
        # 打开保存文件对话框
        target_path = QFileDialog.getSaveFileName(self, "请选择保存路径", '', "(*.txt)")
        # 判断是否选择了文件
        if target_path[0] == '':
            pass
        else:
            # 获取操作日志
            logs = self.plainTextEdit.toPlainText()
            # 获取当前日期时间
            now = datetime.datetime.now()
            # 将操作日志写入文件
            with open(target_path[0], 'w') as f:
                f.write('日志导出时间：' + now.strftime('%Y-%m-%d %H:%M:%S') + '\n')
                f.write(logs)
            QMessageBox.information(self, "提示", "操作日志导出成功！")

    def modify_parameters(self):
        """修改参数"""
        try:
            # 获取当前行行号列号
            row = self.tableWidget.currentRow()
            # 获取当前行的ID
            xx = self.tableWidget.item(row, 7).text()
            yy = self.tableWidget.item(row, 1).text()
            # 将导航页的tabWidget设置为对应的页
            self.show_navigation()
            self.navigation.tabWidget.setCurrentIndex(dict(self.tab_index)[yy])
            # 修改数据中的参数
            self.navigation.modify_judgment = '修改'
            self.navigation.modify_id = xx
        except AttributeError:
            QMessageBox.information(self, "提示", "请先选择一行待修改的数据！")
            pass

    def open_readme(self):
        """打开使用说明"""
        QDesktopServices.openUrl(QUrl('https://gitee.com/fasterthanlight/automatic_clicker'))

    def create_branch(self):
        """创建分支表并重命名"""
        # 弹出输入对话框，提示输入分支名称
        text, ok = QInputDialog.getText(self, "创建分支", "请输入分支名称：")
        if ok:
            try:
                # 连接数据库
                cursor, con = self.sqlitedb()
                # 查找是否有同名分支
                cursor.execute('select 分支表名 from 全局参数 where 分支表名=?', (text,))
                x = cursor.fetchall()
                # print('x:' + str(x))
                # print('x:' + str(len(x)))
                if len(x) > 0:
                    QMessageBox.information(self, "提示", "分支已存在！")
                    return
                else:
                    # 向全局参数表中添加分支表名
                    print('添加分支')
                    cursor.execute('insert into 全局参数(图像文件夹路径,工作簿路径,分支表名,扩展程序) values(?,?,?,?)',
                                   (None, None, text, None))
                    con.commit()
                    # 弹出提示框，提示创建成功
                    QMessageBox.information(self, "提示", "分支创建成功！")
                # 关闭数据库连接
                self.close_database(cursor, con)
                # 加载分支
                self.load_branch()
            except sqlite3.OperationalError:
                QMessageBox.critical(self, "提示", "分支创建失败！")
                pass

    def delete_branch(self):
        """删除分支"""
        # 弹出输入对话框，提示输入分支名称
        print('删除分支')
        cursor, con = self.sqlitedb()
        text = self.comboBox.currentText()
        if text == '主流程':
            QMessageBox.information(self, "提示", "无法删除主分支！")
        else:
            # 将combox显示的名称切换为命令
            self.comboBox.setCurrentText('主流程')
            # 删除分支名称
            cursor.execute('delete from 全局参数 where 分支表名=?', (text,))
            # 关闭数据库连接
            con.commit()
            self.close_database(cursor, con)
            # 将分支名从分支列表中删除
            self.branch_name.remove(text)
            # 弹出提示框
            QMessageBox.information(self, "提示", "分支删除成功！")
            # 重新加载分支列表
            self.load_branch()

    def load_branch(self):
        """加载分支"""
        # 初始化功能
        print('加载分支')
        cursor, con = self.sqlitedb()
        # 获取所有分支名
        cursor.execute("select 分支表名 from 全局参数")
        self.branch_name = [x[0] for x in cursor.fetchall() if x[0] is not None]
        # 关闭数据库连接
        self.close_database(cursor, con)
        self.comboBox.clear()
        self.comboBox.addItems(self.branch_name)


class Setting(QWidget, Ui_Setting):
    """添加设置窗口"""

    def __init__(self):
        super().__init__()
        # 初始化窗体
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 点击保存（应用）按钮
        self.pushButton.clicked.connect(self.save_setting)
        # 点击恢复至默认按钮
        self.pushButton_3.clicked.connect(self.restore_default)
        # 开启极速模式
        self.radioButton_2.clicked.connect(self.speed_mode)
        # 切换普通模式
        self.radioButton.clicked.connect(self.normal_mode)

    def save_setting_date(self):
        """保存设置数据"""
        # 重窗体控件提取数据并放入列表
        list_setting_name = ['图像匹配精度', '时间间隔', '持续时间', '暂停时间', '模式', '启动检查更新']
        image_accuracy = self.horizontalSlider.value() / 10
        interval = self.horizontalSlider_2.value() / 1000
        duration = self.horizontalSlider_3.value() / 1000
        time_sleep = self.horizontalSlider_4.value() / 1000
        model = 1
        if self.checkBox.isChecked():
            update_check = 1
        else:
            update_check = 0
        if self.radioButton_2.isChecked():
            model = 2
        list_setting_value = [image_accuracy, interval, duration, time_sleep, model, update_check]
        # 打开数据库并更新设置数据
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        for i in range(len(list_setting_name)):
            cursor.execute("update 设置 set 值=? where 设置类型=?", (list_setting_value[i], list_setting_name[i]))
            con.commit()
        con.close()

    def save_setting(self):
        """保存按钮事件"""
        self.save_setting_date()
        QMessageBox.information(self, '提醒', '保存成功！')
        self.close()

    def restore_default(self):
        """设置恢复至默认"""
        self.radioButton.isChecked()
        self.horizontalSlider.setValue(9)
        self.horizontalSlider_2.setValue(200)
        self.horizontalSlider_3.setValue(200)
        self.horizontalSlider_4.setValue(100)
        self.save_setting_date()

    def load_setting_data(self):
        """加载设置数据库中的数据"""
        # 连接数据库存入列表
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select * from 设置')
        list_setting_data = cursor.fetchall()
        con.close()
        print(list_setting_data)
        # 设置控件数据为数据库保存的数据
        self.horizontalSlider.setValue(int(list_setting_data[0][1] * 10))
        self.horizontalSlider_2.setValue(int(list_setting_data[1][1] * 1000))
        self.horizontalSlider_3.setValue(int(list_setting_data[2][1] * 1000))
        self.horizontalSlider_4.setValue(int(list_setting_data[3][1] * 1000))
        # 极速模式
        if int(list_setting_data[4][1]) == 2:
            self.radioButton_2.setChecked(True)
            self.pushButton_3.setEnabled(False)
            self.horizontalSlider_2.setEnabled(False)
            self.horizontalSlider_4.setEnabled(False)
        if list_setting_data[5][1] == 1:
            self.checkBox.setChecked(True)
        else:
            self.checkBox.setChecked(False)

    def speed_mode(self):
        """极速模式开启"""
        self.horizontalSlider_2.setValue(0)
        self.horizontalSlider_3.setValue(100)
        self.horizontalSlider_4.setValue(0)
        self.horizontalSlider_2.setEnabled(False)
        self.horizontalSlider_4.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.save_setting_date()

    def normal_mode(self):
        """切换普通模式"""
        self.horizontalSlider_2.setEnabled(True)
        self.horizontalSlider_4.setEnabled(True)
        self.pushButton_3.setEnabled(True)
        self.save_setting_date()


class About(QWidget, Ui_Dialog):
    """关于窗体"""

    def __init__(self):
        super(About, self).__init__()
        self.setupUi(self)
        # 去除窗体最大化、最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)

        self.github.clicked.connect(self.show_github)
        self.gitee.clicked.connect(self.show_gitee)

    def show_github(self):
        # 弹出对话框显示“暂无信息”
        QMessageBox.information(self, '提醒', '暂无信息')
        # QDesktopServices.openUrl(QUrl('https://github.com/FsterThanLight/Clicker'))

    def show_gitee(self):
        QDesktopServices.openUrl(QUrl('https://gitee.com/fasterthanlight/automatic_clicker'))


class Na(QWidget, Ui_navigation):
    """导航页窗体及其功能"""

    def __init__(self, global_window):
        super().__init__()
        # 使用全局变量窗体的一些方法
        self.global_window = global_window
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)
        # 是否激活自定义点击次数
        self.comboBox_3.currentTextChanged.connect(self.spinBox_2_enable)
        # 添加保存按钮事件
        self.modify_judgment = '保存'
        self.modify_id = None
        self.pushButton_2.clicked.connect(lambda: self.save_data(self.modify_judgment, self.modify_id))
        # 获取鼠标位置参数
        self.pushButton_4.clicked.connect(self.mouseMoveEvent)
        # 设置当前日期和时间
        self.checkBox.clicked.connect(self.get_now_date_time)
        # 检查输入的数据是否合法
        self.checkBox_2.clicked.connect(self.check_text_type)
        # 当按钮按下时，获取按键的名称
        self.pushButton_6.clicked.connect(self.print_key_name)
        # 当combobox_8的值改变时，加载combobox的值
        self.comboBox_8.currentTextChanged.connect(lambda: self.find_images(self.comboBox_8, self.comboBox))
        self.comboBox_14.currentTextChanged.connect(lambda: self.find_images(self.comboBox_14, self.comboBox_15))
        self.comboBox_17.currentTextChanged.connect(lambda: self.find_images(self.comboBox_17, self.comboBox_18))
        self.comboBox_12.currentTextChanged.connect(self.find_excel_sheet_name)
        # 切换到导航页时，控制窗口控件的状态
        self.tabWidget.currentChanged.connect(self.tab_widget_change)
        # 调整异常处理选项时，控制窗口控件的状态
        # self.comboBox_9.currentTextChanged.connect(self.exception_handling_judgment_type)
        self.comboBox_9.activated.connect(self.exception_handling_judgment_type)
        # 快捷选择导航页
        self.tab_title = [self.tabWidget.tabText(x) for x in range(self.tabWidget.count())]
        self.comboBox_16.addItems(self.tab_title)
        self.comboBox_16.currentTextChanged.connect(self.quick_select_navigation_page)
        # 行号自动递增提示
        self.checkBox_3.clicked.connect(self.line_number_increasing)
        # 快捷截图功能
        self.pushButton.clicked.connect(lambda: self.quick_screenshot(self.comboBox_8, self.comboBox))
        self.pushButton_7.clicked.connect(lambda: self.delete_all_images(self.comboBox_8, self.comboBox))
        # 信息录入页面的快捷截图功能
        self.pushButton_5.clicked.connect(lambda: self.quick_screenshot(self.comboBox_14, self.comboBox_15))
        self.pushButton_8.clicked.connect(lambda: self.delete_all_images(self.comboBox_14, self.comboBox_15))
        # 网页测试
        # self.pushButton_9.clicked.connect(self.web_functional_testing)

    def load_values_to_controls(self):
        """将值加入到下拉列表中"""
        print('加载导航页下拉列表数据')
        image_folder_path, excel_folder_path, \
            branch_table_name, extenders = self.global_window.extracted_data_global_parameter()
        # 清空下拉列表
        self.comboBox_8.clear()
        self.comboBox_9.clear()
        self.comboBox_12.clear()
        self.comboBox_13.clear()
        self.comboBox_14.clear()
        self.comboBox_11.clear()
        self.comboBox_17.clear()
        self.comboBox_18.clear()
        # 加载下拉列表数据
        self.comboBox_8.addItems(image_folder_path)
        self.comboBox_17.addItems(image_folder_path)
        # 从数据库加载的分支表名
        system_command = ['自动跳过', '抛出异常并暂停', '抛出异常并停止', '扩展程序']
        self.comboBox_9.addItems(system_command)
        self.comboBox_9.addItems(branch_table_name)
        # 从数据库加载的excel表名和图像名称
        self.comboBox_12.addItems(excel_folder_path)
        self.comboBox_14.addItems(image_folder_path)
        # 清空备注
        self.lineEdit_5.clear()

    def quick_select_navigation_page(self):
        """快捷选择导航页"""
        tab_a = self.comboBox_16.currentText()
        tab_index = self.tab_title.index(tab_a)
        self.tabWidget.setCurrentIndex(tab_index)

    def find_images(self, combox, combox_2):
        """选择文件夹并返回文件夹名称"""
        fil_path = combox.currentText()
        try:
            images_name = os.listdir(fil_path)
        except FileNotFoundError:
            images_name = []
        # 去除文件夹中非png文件名称
        for i in range(len(images_name) - 1, -1, -1):
            if ".png" not in images_name[i]:
                images_name.remove(images_name[i])
        print(images_name)
        # 清空combox_2中的所有元素
        combox_2.clear()
        # 将images_name中的所有元素添加到combox_2中
        combox_2.addItems(images_name)
        self.label_3.setText(self.comboBox_8.currentText())

    def find_excel_sheet_name(self):
        """获取excel表格中的所有sheet名称"""
        excel_path = self.comboBox_12.currentText()
        try:
            # 用openpyxl获取excel表格中的所有sheet名称
            excel_sheet_name = openpyxl.load_workbook(excel_path).sheetnames
        except FileNotFoundError:
            excel_sheet_name = []
        except InvalidFileException:
            excel_sheet_name = []
        # 清空combox_13中的所有元素
        self.comboBox_13.clear()
        # 将excel_sheet_name中的所有元素添加到combox_13中
        self.comboBox_13.addItems(excel_sheet_name)

    def print_key_name(self):
        pressed_keys = set()  # create an empty set to store pressed keys
        # # 禁用当前按钮
        self.pushButton_6.setEnabled(False)
        while True:
            event = keyboard.read_event()  # read the keyboard event
            if event.event_type == "down":  # check if the key is pressed down
                pressed_keys.add(event.name)  # add the pressed key to the set
                # 将pressed_keys中的所有元素转换为一行字符串
                pressed_keys_str = list(pressed_keys)
                # pressed_keys_str倒过来
                pressed_keys_str.reverse()
                # 将pressed_keys_str中的所有元素转换为一行字符串
                pressed_keys_str = '+'.join(pressed_keys_str)
                self.label_31.setText(pressed_keys_str)  # print the name of the pressed key
                # print(event.name)  # print the name of the pressed key
            elif event.event_type == "up":  # check if the key is released
                pressed_keys.discard(event.name)  # remove the released key from the set
            if not pressed_keys:  # check if all keys are released
                break  # exit the loop if all keys are released
            # # 激活当前按钮
            self.pushButton_6.setEnabled(True)

    def spinBox_2_enable(self):
        """激活自定义点击次数"""
        if self.comboBox_3.currentText() == '左键（自定义次数）':
            self.spinBox_2.setEnabled(True)
            self.label_22.setEnabled(True)
        else:
            self.spinBox_2.setEnabled(False)
            self.label_22.setEnabled(False)

    def get_mouse_position(self):
        x, y = pyautogui.position()
        self.label_9.setText(str(x))
        self.label_10.setText(str(y))

    def get_now_date_time(self):
        """获取当前日期和时间"""
        now_date_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # 将当前的时间和日期加10分钟
        now_date_time = (datetime.datetime.strptime(now_date_time, '%Y-%m-%d %H:%M:%S') + datetime.timedelta(
            minutes=10)).strftime('%Y-%m-%d %H:%M:%S')
        # 将dateTimeEdit的日期和时间设置为当前日期和时间
        self.dateTimeEdit.setDateTime(datetime.datetime.strptime(now_date_time, '%Y-%m-%d %H:%M:%S'))

    def mouseMoveEvent(self, event):
        # self.setMouseTracking(True)
        self.get_mouse_position()

    def tab_widget_change(self):
        """切换导航页功能"""
        # 获取当前导航页索引
        index = self.tabWidget.currentIndex()
        #     "图像点击": 0,
        #     "坐标点击": 1,
        #     "鼠标移动": 2,
        #     "等待": 3,
        #     "滚轮滑动": 4,
        #     "文本输入": 5,
        #     "按下键盘": 6,
        #     "中键激活": 7,
        #     "鼠标事件": 8,
        #     "excel信息录入": 9
        # 禁用类
        discards = [1, 2, 4, 5, 6, 7, 8]
        discards_not = [0, 3, 9]
        # 不禁用类
        if index in discards:
            self.comboBox_9.setEnabled(True)
            self.comboBox_9.setCurrentIndex(0)
            self.comboBox_9.setEnabled(False)
            self.comboBox_11.setEnabled(True)
            self.comboBox_11.setEnabled(False)
        elif index in discards_not:
            self.comboBox_9.setEnabled(True)
            self.comboBox_11.setEnabled(True)

    def line_number_increasing(self):
        """行号递增功能被选中后弹出提示框"""
        if self.checkBox_3.isChecked():
            QMessageBox.information(self, '提示',
                                    '启用该功能后，请在主页面中设置循环次数大于1，执行全部指令后，循环执行时，单元格行号会自动递增。',
                                    QMessageBox.Ok)

    def exception_handling_judgment(self):
        """判断异常处理方式"""
        exception_handling_text = None

        def remove_none(list_):
            """去除列表中的none"""
            list_x = []
            for i in list_:
                if i[0] is not None:
                    list_x.append(i[0])
            return list_x

        if self.comboBox_9.currentText() == '自动跳过':
            exception_handling_text = '自动跳过'
        elif self.comboBox_9.currentText() == '抛出异常并暂停':
            exception_handling_text = '抛出异常并暂停'
        elif self.comboBox_9.currentText() == '抛出异常并停止':
            exception_handling_text = '抛出异常并停止'
        elif self.comboBox_9.currentText() == '扩展程序':
            exception_handling_text = self.comboBox_11.currentText()
        else:
            # 连接数据库
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            # 获取表中数据记录的个数
            cursor.execute('SELECT 分支表名 FROM 全局参数')
            result = cursor.fetchall()
            branch_table_name = remove_none(result)
            # print(branch_table_name)
            cursor.close()
            con.close()
            branch_table_name_index = branch_table_name.index(self.comboBox_9.currentText())
            exception_handling_text = '分支-' + str(branch_table_name_index) + '-' + str(
                int(self.comboBox_10.currentText()) - 1)
        # print('异常处理方式：', exception_handling_text)
        return exception_handling_text

    def exception_handling_judgment_type(self):
        """判断异常护理选项并调整控件"""
        system_command = ['自动跳过', '抛出异常并暂停', '抛出异常并停止']
        try:
            if self.comboBox_9.currentText() in system_command:
                # 开始位置
                self.comboBox_10.clear()
                self.comboBox_10.setEnabled(False)
                # 扩展程序
                self.comboBox_11.clear()
                self.comboBox_11.setEnabled(False)
            elif self.comboBox_9.currentText() not in system_command and self.comboBox_9.currentText() != '扩展程序':
                # 扩展程序
                self.comboBox_11.clear()
                self.comboBox_11.setEnabled(False)
                self.comboBox_10.setEnabled(True)
                # 连接数据库
                con = sqlite3.connect('命令集.db')
                cursor = con.cursor()
                # 获取表中数据记录的个数
                branch_name = self.comboBox_9.currentText()
                cursor.execute('SELECT count(*) FROM 命令 where 隶属分支=?', (branch_name,))
                count_record = cursor.fetchone()[0]
                # 关闭连接
                cursor.close()
                con.close()
                self.comboBox_10.clear()
                # 加载分支中的命令序号
                branch_order = [str(i) for i in range(1, count_record + 1)]
                if len(branch_order) == 0:
                    # 弹出警告框
                    self.comboBox_9.setCurrentIndex(0)
                    QMessageBox.warning(self, '警告', '该分支下没有指令，请先添加！', QMessageBox.Yes)
                else:
                    self.comboBox_10.addItems(branch_order)

            elif self.comboBox_9.currentText() not in system_command and self.comboBox_9.currentText() == '扩展程序':
                # 开始位置
                self.comboBox_10.clear()
                self.comboBox_10.setEnabled(False)
                # 扩展程序
                self.comboBox_11.setEnabled(True)
                image_folder_path, excel_folder_path, \
                    branch_table_name, extenders = self.global_window.extracted_data_global_parameter()
                self.comboBox_11.clear()
                self.comboBox_11.addItems(extenders)
        except sqlite3.OperationalError:
            pass

    def check_text_type(self):
        """检查文本输入类型"""
        text = self.textEdit.toPlainText()
        # 检查text中是否为英文大小写字母和数字
        if re.search('[a-zA-Z0-9]', text) is None:
            self.checkBox_2.setChecked(False)
            QMessageBox.warning(self, '警告', '文本输入仅支持输入英文大小写字母和数字！', QMessageBox.Yes)

    def quick_screenshot(self, combox, combox_2):
        """截图功能"""
        if combox.currentText() == '':
            QMessageBox.warning(self, '警告', '未选择图像文件夹！', QMessageBox.Yes)
        else:
            # 隐藏主窗口
            self.hide()
            main_window.hide()
            # 截图
            screen_capture = ScreenCapture()
            screen_capture.screenshot_area()
            # 显示主窗口
            self.show()
            main_window.show()
            # 文件夹路径和文件名
            image_folder_path = combox.currentText()
            image_name, ok = QInputDialog.getText(self, "截图", "请输入图像名称：")
            if ok:
                # 检查image_name是否包含中文字符
                if re.search('[\u4e00-\u9fa5]', image_name) is not None:
                    QMessageBox.warning(self, '警告', '图像名称暂不支持中文字符！保存失败。', QMessageBox.Yes)
                else:
                    screen_capture.screen_shot(image_folder_path, image_name)
            # 刷新图像文件夹
            self.find_images(combox, combox_2)
            main_window.plainTextEdit.appendPlainText('已快捷截图：' + image_name)
            combox_2.setCurrentText(image_name)

    def delete_all_images(self, combox, combox_2):
        if combox.currentText() == '':
            pass
        else:
            file_path = combox.currentText()
            # 删除文件夹中所有文件，保留文件夹
            shutil.rmtree(file_path)
            os.mkdir(file_path)
            self.find_images(combox, combox_2)
            # 弹出提示框
            QMessageBox.information(self, '提示', '已删除所有图像！', QMessageBox.Yes)

    def save_data(self, judge='保存', xx=None):
        """获取4个参数命令，并保存至数据库"""

        def writes_commands_to_the_database(instruction, repeat_number, exception_handling, image=None,
                                            parameter_1=None,
                                            parameter_2=None, parameter_3=None, parameter_4=None, remarks=None):
            """向数据库写入命令"""
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            branch_name = main_window.comboBox.currentText()
            try:
                if judge == '保存':
                    cursor.execute(
                        'INSERT INTO 命令(图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) VALUES (?,?,?,?,?,?,?,?,?,?)',
                        (image, instruction, parameter_1, parameter_2, parameter_3, parameter_4, repeat_number,
                         exception_handling, remarks, branch_name))
                elif judge == '修改':
                    cursor.execute(
                        'UPDATE 命令 SET 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=?,备注=? WHERE ID=?',
                        (image, instruction, parameter_1, parameter_2, parameter_3, parameter_4, repeat_number,
                         exception_handling, remarks, xx))
                con.commit()
                con.close()
            except sqlite3.OperationalError:
                QMessageBox.critical(self, "错误", "无写入数据权限，请以管理员身份运行！")

        def time_judgment(target_time):
            """判断时间是否大于当前时间"""
            # 获取当前时间年月日和时分秒
            # print(target_time)
            now_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
            # 将now_time转换为时间格式
            now_time = datetime.datetime.strptime(now_time, '%Y/%m/%d %H:%M:%S')
            # 将字符参数转换为时间格式
            target_time = datetime.datetime.strptime(target_time, '%Y/%m/%d %H:%M:%S')
            # 判断是否重新输入
            # print(now_time)
            # print(target_time)
            xx = 0
            if now_time < target_time:
                print('目标时间大于当前时间，正确')
                xx = 0
            else:
                print('目标时间小于当前时间，错误')
                xx = 1
            return xx

        # 判断当前tab页
        # 读取功能区的参数：重复次数、异常处理、备注
        repeat_number = self.spinBox.value()
        exception_handling = self.exception_handling_judgment()
        remarks = self.lineEdit_5.text()
        # 图像点击事件的参数获取
        if self.tabWidget.currentIndex() == 0:
            # 获取5个参数命令，写入数据库
            instruction = "图像点击"
            image = self.comboBox_8.currentText() + '/' + self.comboBox.currentText()
            parameter_1 = self.comboBox_2.currentText()
            # 如果复选框被选中，则获取第二个参数
            parameter_2 = None
            if self.radioButton_2.isChecked():
                parameter_2 = '自动略过'
            elif self.radioButton_4.isChecked():
                parameter_2 = self.spinBox_4.value()
            # 将命令写入数据库
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            image=image, parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)
            print('已经保存图像识别点击的数据至数据库')
        # 鼠标点击事件的参数获取
        elif self.tabWidget.currentIndex() == 1:
            instruction = "坐标点击"
            parameter_1 = self.comboBox_3.currentText()
            parameter_2 = self.label_9.text() + "-" + self.label_10.text() + "-" + str(self.spinBox_2.value())
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)

        elif self.tabWidget.currentIndex() == 2:
            # 获取5个参数命令
            instruction = "鼠标移动"
            # 获取鼠标移动的参数
            # 鼠标移动的方向
            parameter_1 = self.comboBox_4.currentText()
            # 鼠标移动的距离
            try:
                parameter_2 = int(self.lineEdit.text())
            except ValueError:
                QMessageBox.critical(self, "错误", "移动距离请输入数字！")
                return
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)
            print('已经保存鼠标移动的数据至数据库')
        # 等待事件的参数获取
        elif self.tabWidget.currentIndex() == 3:
            instruction = "等待"
            # 获取等待的参数
            # 如果checkBox没有被选中，则第一个参数为等待时间
            image = None
            parameter_1 = None
            parameter_2 = None
            parameter_3 = None
            if not self.checkBox.isChecked() and not self.checkBox_5.isChecked():
                parameter_1 = "等待"
                try:
                    parameter_2 = int(self.lineEdit_2.text())
                except ValueError:
                    QMessageBox.critical(self, "错误", "等待时间请输入数字！")
                    return
            elif self.checkBox.isChecked() and not self.checkBox_5.isChecked():
                parameter_1 = "等待到指定时间"
                # 判断时间是否大于当前时间
                parameter_2 = self.dateTimeEdit.text() + "+" + self.comboBox_6.currentText()
                try:
                    xx = time_judgment(parameter_2.split('+')[0])
                    if xx == 1:
                        raise TimeoutError("Invalid number!")
                except TimeoutError:
                    QMessageBox.critical(self, "错误", "启动时间小于当前系统时间，无效的指令。")
                    return
            elif not self.checkBox.isChecked() and self.checkBox_5.isChecked():
                parameter_1 = "等待到指定图片"
                image = self.comboBox_17.currentText() + '/' + self.comboBox_18.currentText()
                parameter_2 = self.comboBox_19.currentText()
                parameter_3 = self.spinBox_6.value()
            elif self.checkBox.isChecked() and self.checkBox_5.isChecked():
                QMessageBox.critical(self, "错误", "等待指定时间和等待指定图片不能同时勾选！")
                return
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            image=image,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2,
                                            parameter_3=parameter_3, remarks=remarks)

        # 鼠标滚轮滑动事件的参数获取
        elif self.tabWidget.currentIndex() == 4:
            # 获取5个参数命令
            instruction = "滚轮滑动"
            # 获取鼠标滚轮滑动的参数
            # 鼠标滚轮滑动的方向
            parameter_1 = self.comboBox_5.currentText()
            # 鼠标滚轮滑动的距离
            try:
                parameter_2 = self.lineEdit_3.text()
            except ValueError:
                QMessageBox.critical(self, "错误", "滑动距离请输入数字！")
                return
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)
            print('已经保存鼠标滚轮滑动的数据至数据库')
        # 文本输入事件的参数获取
        elif self.tabWidget.currentIndex() == 5:
            # 获取5个参数命令
            instruction = "文本输入"
            # 获取文本输入的参数
            # 文本输入的内容
            parameter_1 = self.textEdit.toPlainText()
            parameter_2 = str(self.checkBox_2.isChecked())
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)
            print('已经保存文本输入的数据至数据库')
        # 按下键盘事件的参数获取
        elif self.tabWidget.currentIndex() == 6:
            instruction = "按下键盘"
            # 获取按下键盘的参数
            # 按下键盘的内容
            parameter_1 = self.label_31.text()
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1, remarks=remarks)
            print('已经保存按键的数据至数据库')
        # 中键激活事件的参数获取
        elif self.tabWidget.currentIndex() == 7:
            instruction = "中键激活"
            # 获取中键激活的参数
            # 中键激活的内容
            parameter_1 = None
            parameter_2 = None
            if self.radioButton.isChecked():
                parameter_1 = '模拟点击'
                parameter_2 = self.spinBox_3.value()
            elif self.radioButton_2.isChecked():
                parameter_1 = '自定义'
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2, remarks=remarks)
        # 鼠标当前位置事件的参数获取
        elif self.tabWidget.currentIndex() == 8:
            instruction = "鼠标事件"
            # 获取鼠标当前位置的参数
            parameter_1 = self.comboBox_7.currentText()
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1, remarks=remarks)
        # excel信息录入功能的参数获取
        elif self.tabWidget.currentIndex() == 9:
            instruction = "excel信息录入"
            parameter_4 = None
            # 获取excel工作簿路径和工作表名称
            parameter_1 = self.comboBox_12.currentText() + "-" + self.comboBox_13.currentText()
            # 获取图像文件路径
            image = self.comboBox_14.currentText() + '/' + self.comboBox_15.currentText()
            # 获取单元格值
            parameter_2 = self.lineEdit_4.text()
            # 判断是否递增行号和特殊控件输入
            parameter_3 = str(self.checkBox_3.isChecked()) + '-' + str(self.checkBox_4.isChecked())
            # 判断其他参数
            if self.radioButton_3.isChecked() and not self.radioButton_5.isChecked():
                parameter_4 = '自动跳过'
            elif not self.radioButton_3.isChecked() and self.radioButton_5.isChecked():
                parameter_4 = self.spinBox_5.value()

            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2,
                                            parameter_3=parameter_3,
                                            parameter_4=parameter_4,
                                            image=image, remarks=remarks)
        # 网页操作功能的参数获取
        elif self.tabWidget.currentIndex() == 10:
            instruction = "打开网址"
            # 获取网页链接
            web_page_link = self.lineEdit_6.text()
            # 写入数据库
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            image=web_page_link, remarks=remarks)

        # 关闭窗体
        self.close()
        self.modify_judgment = '保存'
        self.modify_id = None


class Info(QDialog, Ui_Form):
    def __init__(self, parent=None):
        super(Info, self).__init__(parent)
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)


class Global_s(QDialog, Ui_Global):
    """全局参数设置窗体"""

    def __init__(self, parent=None):
        super(Global_s, self).__init__(parent)
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)
        # 刷新listview
        self.refresh_listview()
        # 添加图像文件夹路径
        self.pushButton.clicked.connect(lambda: self.select_file("图像文件夹路径"))
        # 添加工作簿路径
        self.pushButton_3.clicked.connect(lambda: self.select_file("工作簿路径"))
        # 添加扩展程序
        self.pushButton_9.clicked.connect(lambda: self.select_file("扩展程序"))
        # 删除listview中的项
        self.pushButton_2.clicked.connect(lambda: self.delete_listview(self.listView, "图像文件夹路径"))
        self.pushButton_4.clicked.connect(lambda: self.delete_listview(self.listView_2, "工作簿路径"))
        self.pushButton_10.clicked.connect(lambda: self.delete_listview(self.listView_5, "扩展程序"))

    def select_file(self, judge):
        """选择文件"""
        if judge == "图像文件夹路径":
            fil_path = QFileDialog.getExistingDirectory(self, "选择存储目标图像的文件夹")
            if fil_path != '':
                self.write_to_database(fil_path, None, None, None)
        elif judge == "工作簿路径":
            fil_path, _ = QFileDialog.getOpenFileName(self, "选择工作簿", filter="Excel 工作簿(*.xlsx)")
            if fil_path != '':
                self.write_to_database(None, fil_path, None, None)
        elif judge == "扩展程序":
            # 打开文件对话框，选择一个py或exe文件
            fil_path, _ = QFileDialog.getOpenFileName(self, "选择扩展程序",
                                                      filter="Python文件(*.py);;可执行文件(*.exe)")
            if fil_path != '':
                self.write_to_database(None, None, None, fil_path)
        self.refresh_listview()

    def delete_listview(self, list_view, judge):
        """删除listview中选中的那行数据"""
        # 获取选中的行的值
        try:
            indexes = list_view.selectedIndexes()
            item = list_view.model().itemFromIndex(indexes[0])
            value = item.text()
            print("删除的值为：", value)
            # 删除数据库中的数据
            self.delete_data(value, judge)
            # 刷新listview
            self.refresh_listview()
        except AttributeError:
            pass
        except IndexError:
            pass

    def delete_data(self, value, judge):
        """删除数据库中的数据"""
        # 连接数据库
        conn = sqlite3.connect('命令集.db')
        c = conn.cursor()
        # 删除数据
        if judge == '图像文件夹路径':
            c.execute("DELETE FROM 全局参数 WHERE 图像文件夹路径 = ?", (value,))
        elif judge == '工作簿路径':
            c.execute("DELETE FROM 全局参数 WHERE 工作簿路径 = ?", (value,))
        elif judge == '扩展程序':
            c.execute("DELETE FROM 全局参数 WHERE 扩展程序 = ?", (value,))
        # 删除无用数据
        c.execute("DELETE FROM 全局参数 WHERE 图像文件夹路径 is NULL and "
                  "工作簿路径 is NULL and 分支表名 is NULL and 扩展程序 is NULL")
        conn.commit()
        conn.close()

    def refresh_listview(self):
        """刷新listview"""
        # 获取数据库中的数据
        image_folder_path, excel_folder_path, \
            branch_table_name, extenders = self.extracted_data_global_parameter()

        print('扩展程序：', extenders)

        def add_listview(list_, listview):
            """添加listview"""
            listview.setModel(QStandardItemModel())
            if len(list_) != 0:
                model_2 = listview.model()
                for i in list_:
                    model_2.appendRow(QStandardItem(i))

        add_listview(image_folder_path, self.listView)
        add_listview(excel_folder_path, self.listView_2)
        add_listview(extenders, self.listView_5)

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
        return image_folder_path, excel_folder_path, branch_table_name, extenders

    def write_to_database(self, images_file, work_book_path, branch_table_name, extension_program):
        """将全局参数写入数据库"""
        # 连接数据库
        conn = sqlite3.connect('命令集.db')
        c = conn.cursor()
        # 向数据库中的“图像文件夹路径”字段添加文件夹路径
        c.execute('INSERT INTO 全局参数(图像文件夹路径,工作簿路径,分支表名,扩展程序) VALUES (?,?,?,?)',
                  (images_file, work_book_path, branch_table_name, extension_program))
        conn.commit()
        conn.close()


if __name__ == "__main__":
    # 自适应高分辨率
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)

    app = QApplication([])
    # 创建主窗体
    main_window = Main_window()
    # 显示窗体，并根据设置检查更新
    main_window.main_show()
    # 显示添加对话框窗口
    sys.exit(app.exec_())

    # def is_admin():
    #     try:
    #         return ctypes.windll.shell32.IsUserAnAdmin()
    #     except:
    #         return False
    #
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
