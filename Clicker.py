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

import datetime
import json
import os
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

from main_work import MainWork, exit_main_work
from 窗体.about import Ui_Dialog
from 窗体.global_s import Ui_Global
from 窗体.info import Ui_Form
from 窗体.mainwindow import Ui_MainWindow
from 窗体.navigation import Ui_navigation
from 窗体.setting import Ui_Setting

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'}
# 声明一个路径的全局变量
fil_path = ''


def load_json():
    """从json文件中加载更新网址和保留文件名"""
    file_name = 'update_data.json'
    with open(file_name, 'r', encoding='utf8') as f:
        data = json.load(f)
    url = cryptocode.decrypt(data['url_encrypt'], '123456')
    # list_keep = []
    # for v in data.values():
    #     list_keep.append(v)
    print(url)
    # print(list_keep)
    return url


def get_download_address(main_window, warning):
    """获取下载地址、版本信息、更新说明"""
    global headers
    url = load_json()
    print(url)
    try:
        res = requests.get(url, headers=headers, timeout=0.2)
        info = cryptocode.decrypt(res.text, '123456')
        list_1 = info.split('=')
        print(list_1)
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
        # 实例化子窗口1
        # self.dialog_1 = Dialog()
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
        # # 单元格变动自动存储
        # self.change_state = True
        # self.tableWidget.cellChanged.connect(lambda: self.table_cell_changed(False))
        # 保存按钮
        self.actionb.triggered.connect(self.save_data_to_current)
        # 清空指令按钮
        self.toolButton_6.clicked.connect(self.clear_table)
        # 导入数据按钮
        # self.actionf.triggered.connect(self.data_import)
        # 主窗体开始按钮
        self.pushButton_5.clicked.connect(self.start)
        # 实时计时
        # self.lcd_time = 1
        # self.timer = QTimer()
        # self.timer.timeout.connect(lambda: self.display_running_time('显示时间'))
        # 打开设置
        self.actions_2.triggered.connect(self.show_setting)
        # 结束任务按钮
        self.pushButton_6.clicked.connect(exit_main_work)
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
        self.pushButton_8.clicked.connect(lambda: self.modify_parameters(self.tab_index))

    # def show_dialog(self):
    #     self.dialog_1.show()
    #     print('子窗口开启')
    #     resize = self.geometry()
    #     self.dialog_1.move(resize.x() + 50, resize.y() + 200)

    def format_table(self):
        """设置主窗口表格格式"""
        # 列的大小拉伸，可被调整
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 列的大小为可交互式的，用户可以调整
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.tableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.tableWidget.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        # # 列的大小调整为固定，列宽不会改变
        self.tableWidget.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed)
        self.tableWidget.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
        # 设置列宽为50像素
        self.tableWidget.setColumnWidth(7, 30)
        self.tableWidget.setColumnWidth(8, 30)

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
        print('获取数据')
        print(fil_path)
        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(0)
        # 进度条归零
        self.progressBar.setValue(0)
        # 获取数据库数据
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select 图像名称,指令类型,异常处理,参数1,参数2,参数3,参数4,重复次数,ID from 命令')
        list_order = cursor.fetchall()
        con.close()
        # 在表格中写入数据
        print(list_order)
        for i in range(len(list_order)):
            self.tableWidget.insertRow(i)
            for j in range(len(list_order[i])):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(list_order[i][j])))

    def delete_data(self):
        """删除选中的数据行"""
        # 获取选中值的行号和id
        row = self.tableWidget.currentRow()
        column = self.tableWidget.currentColumn()
        xx = self.tableWidget.item(row, 8).text()
        print(row, column, xx)
        # 删除数据库中指定id的数据
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('delete from 命令 where ID=?', (xx,))
        con.commit()
        con.close()
        # 调用get_data()函数，刷新表格
        self.get_data()

        # if row != -1:
        #     xx = self.tableWidget.item(row, 4).text()
        # else:
        #     xx = -1
        # try:
        #     self.tableWidget.removeRow(row)
        #     # 删除数据库中数据
        #     con = sqlite3.connect('命令集.db')
        #     cursor = con.cursor()
        #     cursor.execute('delete from 命令 where ID=?', (xx,))
        #     con.commit()
        #     con.close()
        # except UnboundLocalError:
        #     pass

    def go_up_down(self, judge):
        """向上或向下移动选中的行"""
        # 获取选中值的行号和id
        row = self.tableWidget.currentRow()
        column = self.tableWidget.currentColumn()
        try:
            xx = self.tableWidget.item(row, 8).text()
            # 将选中行的数据在数据库中与上一行数据交换，如果是第一行则不交换
            id = int(self.tableWidget.item(row, 8).text())
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
                cursor.execute(
                    'select 图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,ID from 命令 where ID=?', (id,))
                list_id = cursor.fetchall()
                cursor.execute(
                    'select 图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,ID from 命令 where ID=?',
                    (id_up_down,))
                list_id_up = cursor.fetchall()
                # 交换选中行和上一行的数据
                cursor.execute(
                    'update 命令 set 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=? where ID=?',
                    (
                        list_id_up[0][0], list_id_up[0][1], list_id_up[0][2], list_id_up[0][3], list_id_up[0][4],
                        list_id_up[0][5], list_id_up[0][6], list_id_up[0][7], id))
                cursor.execute(
                    'update 命令 set 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=? where ID=?',
                    (list_id[0][0], list_id[0][1], list_id[0][2], list_id[0][3], list_id[0][4], list_id[0][5],
                     list_id[0][6], list_id[0][7], id_up_down))
                con.commit()
                con.close()
            # 调用get_data()函数，刷新表格
            self.get_data()
            # 将焦点移动到交换后的行
            self.tableWidget.setCurrentCell(row_up_down, column)
        except AttributeError:
            pass

        # def table_cell_changed(self, combox_change):
        #     """单元格改变时自动存储"""
        #     if self.change_state:
        #         print('自动存储')
        #         row = self.tableWidget.currentRow()
        #         if combox_change:
        #             self.tableWidget.item(row, 2).setText('0')
        #         else:
        #             pass
        #         # 获取选中行的id，及其他参数
        #         id = self.tableWidget.item(row, 4).text()
        #         images = self.tableWidget.item(row, 0).text()
        #         parameter = self.tableWidget.item(row, 2).text()
        #         repeat_number = self.tableWidget.item(row, 3).text()
        #         option = self.tableWidget.cellWidget(row, 1).currentText()
        #         # 连接数据库，提交修改
        #         con = sqlite3.connect('命令集.db')
        #         cursor = con.cursor()
        #         cursor.execute('update 命令 set 图像名称=?,键鼠命令=?,参数=?,重复次数=? where ID=?',
        #                        (images, option, parameter, repeat_number, id))
        #         con.commit()
        #         con.close()

    def save_data_to_current(self):
        """保存配置文件到当前文件夹下"""
        # 获取图像文件夹路径
        global fil_path
        # 提取数据库中的所有数据
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select * from 命令')
        list_orders = cursor.fetchall()
        con.close()
        # 判断是否选择了文件夹
        if fil_path == '':
            fil_path = QFileDialog.getExistingDirectory(self, "选择保存路径。")
        elif fil_path != '':
            pass
        # 创建txt文件
        file = fil_path + "/命令集.txt"
        # 向txt中写入数据
        try:
            with open(file, 'w', encoding='utf-8') as f:
                f.write('请将本文件放入保存图像的文件夹中。\n')
                for i in range(len(list_orders)):
                    for j in range(len(list_orders[i])):
                        f.write(str(list_orders[i][j]) + ',')
                    f.write('\n')
                QMessageBox.information(self, '保存成功', '数据已保存至' + file)
        except PermissionError:
            QMessageBox.warning(self, '保存失败', '无效的文件路径。')

    # QMessageBox.warning(self, '未选择文件夹', "请点击'添加指令'并选择存放目标图像的文件夹！")

    def clear_database(self):
        """清空数据库"""
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('delete from 命令 where ID<>-1')
        con.commit()
        con.close()

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
        global fil_path
        data_file_path = QFileDialog.getOpenFileName(self, "请选择'命令集.txt'", '', "(*.txt)")
        print(data_file_path)
        # 获取命令集文件夹路径
        fil_path = '/'.join(data_file_path[0].split('/')[0:-1])
        self.navigation.select_file(1)
        # 清空数据库并导入新数据
        # if data_file_path[0] != '':
        self.clear_database()
        with open(data_file_path[0], 'r', encoding='utf-8') as f:
            list_order = f.readlines()
            print(list_order)
            for i in list_order:
                j = i.split(',')
                if len(j) == 7:
                    # 将txt文本转化为数据库对应参数
                    id = int(j[0])
                    image_name = j[1]
                    instruction = j[2]
                    parameter = j[3]
                    parameter_2 = j[4]
                    repeat_number = int(j[5])
                    # 连接数据库，插入数据
                    con = sqlite3.connect('命令集.db')
                    cursor = con.cursor()
                    try:
                        print("插入数据")
                        cursor.execute('insert into 命令(ID,图像名称,键鼠命令,参数,参数2,重复次数) values(?,?,?,?,?,?)',
                                       (id, image_name, instruction, parameter, parameter_2, repeat_number))
                    except sqlite3.IntegrityError:
                        pass
                    con.commit()
                    con.close()
            self.get_data()

    def start(self):
        """主窗体开始按钮"""
        global fil_path
        # mainWork(fil_path, self)
        self.info.show()
        # print("导航页窗口开启")
        resize = self.geometry()
        self.info.move(resize.x() + 45, resize.y() - 30)
        # 开始主任务
        self.main_work.file_path = fil_path
        self.main_work.start_work()
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

    def modify_parameters(self, tab_index):
        """修改参数"""
        self.show_navigation()
        # 获取当前行行号列号
        row = self.tableWidget.currentRow()
        # 获取当前行的ID
        try:
            xx = self.tableWidget.item(row, 8).text()
            yy = self.tableWidget.item(row, 1).text()
            # 将导航页的tabWidget设置为对应的页
            self.navigation.tabWidget.setCurrentIndex(dict(tab_index)[yy])
            # 修改数据中的参数
            self.navigation.modify_judgment = '修改'
            self.navigation.modify_id = xx
        except AttributeError:
            pass

    def open_readme(self):
        """打开使用说明"""
        QDesktopServices.openUrl(QUrl('https://gitee.com/fasterthanlight/automatic_clicker'))


# class Dialog(QWidget, Ui_Form):
#     """添加指令对话框"""
#
#     def __init__(self):
#         super().__init__()
#         # 初始化窗体
#         self.setupUi(self)
#         self.pushButton_3.clicked.connect(lambda: self.select_file(0))
#         self.spinBox_2.setValue(1)
#         self.pushButton.clicked.connect(self.save_data)
#         self.filePath = ''
#         # 设置子窗口出现阻塞主窗口
#         self.setWindowModality(Qt.ApplicationModal)
#         self.list_combox_3_value = []
#         list_controls = [self.textEdit, self.spinBox, self.spinBox_2, self.comboBox,
#                          self.comboBox_3]
#         for i in list_controls:
#             i.setEnabled(False)
#
#     def select_file(self, judge):
#         """选择文件夹并返回文件夹名称"""
#         if judge == 0:
#             self.filePath = QFileDialog.getExistingDirectory(self, "选择存储目标图像的文件夹")
#         try:
#             images_name = os.listdir(self.filePath)
#         except FileNotFoundError:
#             images_name = []
#         # 去除文件夹中非png文件名称
#         for i in range(len(images_name) - 1, -1, -1):
#             if ".png" not in images_name[i]:
#                 images_name.remove(images_name[i])
#         print(images_name)
#         self.label_6.setText(self.filePath.split('/')[-1])
#         self.comboBox.addItems(images_name)
#         self.comboBox_2.currentIndexChanged.connect(self.change_label3)
#         self.comboBox.setEnabled(True)
#         self.spinBox_2.setEnabled(True)
#
#     def change_label3(self):
#         """标签3根据下拉框2的选择变化"""
#         self.spinBox_2.setValue(1)
#         combox_text = self.comboBox_2.currentText()
#
#         def commonly_used_controls(dialog_1):
#             """常用控件恢复运行"""
#             dialog_1.label_2.setStyleSheet('color:red')
#             dialog_1.comboBox.setEnabled(True)
#             dialog_1.spinBox_2.setEnabled(True)
#             dialog_1.label_4.setStyleSheet('color:red')
#
#         def all_disabled(dialog_1):
#             """指令框所有控件全部禁用"""
#             list_controls = [dialog_1.textEdit, dialog_1.spinBox, dialog_1.spinBox_2, dialog_1.comboBox,
#                              dialog_1.comboBox_3]
#             list_label = [dialog_1.label_2, dialog_1.label_3, dialog_1.label_4, dialog_1.label_7,
#                           dialog_1.label_8]
#             for i in list_controls:
#                 i.setEnabled(False)
#             for i in list_label:
#                 i.setStyleSheet('color:transparent')
#             dialog_1.comboBox_3.clear()
#
#         if combox_text == '等待':
#             all_disabled(self)
#             commonly_used_controls(self)
#             self.label_3.setStyleSheet('color:red')
#             self.spinBox.setEnabled(True)
#             self.label_3.setText('等待时长')
#
#         if combox_text == '左键单击':
#             all_disabled(self)
#             commonly_used_controls(self)
#
#         if combox_text == '左键双击':
#             all_disabled(self)
#             commonly_used_controls(self)
#
#         if combox_text == '右键单击':
#             all_disabled(self)
#             commonly_used_controls(self)
#
#         if combox_text == '滚轮滑动':
#             all_disabled(self)
#             commonly_used_controls(self)
#             self.label_3.setStyleSheet('color:red')
#             self.label_3.setText('滑动距离')
#             self.label_8.setStyleSheet('color:red')
#             self.label_8.setText('滑动方向')
#             self.list_combox_3_value = ['向上滑动', '向下滑动']
#             self.comboBox_3.addItems(self.list_combox_3_value)
#             self.comboBox_3.setEnabled(True)
#             self.spinBox.setEnabled(True)
#
#         if combox_text == '内容输入':
#             all_disabled(self)
#             commonly_used_controls(self)
#             self.label_7.setStyleSheet('color:red')
#             self.textEdit.setEnabled(True)
#
#         if combox_text == '鼠标移动':
#             all_disabled(self)
#             commonly_used_controls(self)
#             self.label_8.setStyleSheet('color:red')
#             self.label_8.setText('移动方向')
#             self.label_3.setStyleSheet('color:red')
#             self.label_3.setText('移动距离')
#             self.list_combox_3_value = ['向上', '向下', '向左', '向右']
#             self.comboBox_3.addItems(self.list_combox_3_value)
#             self.comboBox_3.setEnabled(True)
#             self.spinBox.setEnabled(True)
#
#     def save_data(self):
#         """获取4个参数命令，并保存至数据库"""
#         instruction = self.comboBox_2.currentText()
#         # 根据参数的不同获取不同位置的4个参数
#         # 获取图像名称和重读次数
#         image = self.comboBox.currentText()
#         repeat_number = self.spinBox_2.value()
#         parameter = ''
#         # 获取鼠标单击事件或等待的参数
#         list_click = ['左键单击', '左键双击', '右键单击', '等待']
#         if instruction in list_click:
#             parameter = self.spinBox.value()
#         # 获取滚轮滑动事件参数
#         if instruction == '滚轮滑动':
#             direction = self.comboBox_3.currentText()
#             if direction == '向上滑动':
#                 parameter = self.spinBox.value()
#             elif direction == '向下滑动':
#                 x = int(self.spinBox.value())
#                 parameter = str(x - 2 * x)
#         # 获取内容输入事件的参数
#         if instruction == '内容输入':
#             parameter = self.textEdit.toPlainText()
#         # 获取鼠标移动的事件参数
#         if instruction == '鼠标移动':
#             direction = self.comboBox_3.currentText()
#             distance = self.spinBox.value()
#             parameter = direction + '-' + str(distance)
#         # 连接数据库，将数据插入表中并关闭数据库
#         if self.filePath != '':
#             con = sqlite3.connect('命令集.db')
#             cursor = con.cursor()
#             try:
#                 cursor.execute('INSERT INTO 命令(图像名称,键鼠命令,参数,重复次数) VALUES (?,?,?,?)',
#                                (image, instruction, parameter, repeat_number))
#                 con.commit()
#                 con.close()
#             except sqlite3.OperationalError:
#                 QMessageBox.critical(self, "错误", "无写入数据权限，请以管理员身份运行！")
#         self.close()


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
        # 当按钮按下时，获取按键的名称
        self.pushButton_6.clicked.connect(self.print_key_name)
        # 当combobox_8的值改变时，加载combobox的值
        self.comboBox_8.currentTextChanged.connect(lambda: self.find_images(self.comboBox_8, self.comboBox))
        self.comboBox_14.currentTextChanged.connect(lambda: self.find_images(self.comboBox_14, self.comboBox_15))
        self.comboBox_12.currentTextChanged.connect(self.find_excel_sheet_name)

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
        # 加载下拉列表数据
        self.comboBox_8.addItems(image_folder_path)
        self.comboBox_9.addItems(branch_table_name)
        self.comboBox_12.addItems(excel_folder_path)
        self.comboBox_14.addItems(image_folder_path)
        self.comboBox_11.addItems(extenders)

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
        # 将dateTimeEdit的日期和时间设置为当前日期和时间
        self.dateTimeEdit.setDateTime(datetime.datetime.strptime(now_date_time, '%Y-%m-%d %H:%M:%S'))

    def mouseMoveEvent(self, event):
        # self.setMouseTracking(True)
        self.get_mouse_position()

    def save_data(self, judge='保存', xx=None):
        """获取4个参数命令，并保存至数据库"""

        def writes_commands_to_the_database(instruction, repeat_number, exception_handling, image=None,
                                            parameter_1=None,
                                            parameter_2=None, parameter_3=None, parameter_4=None):
            """向数据库写入命令"""
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            try:
                if judge == '保存':
                    cursor.execute(
                        'INSERT INTO 命令(图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理) VALUES (?,?,?,?,?,?,?,?)',
                        (image, instruction, parameter_1, parameter_2, parameter_3, parameter_4, repeat_number,
                         exception_handling))
                elif judge == '修改':
                    cursor.execute(
                        'UPDATE 命令 SET 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=? WHERE ID=?',
                        (image, instruction, parameter_1, parameter_2, parameter_3, parameter_4, repeat_number,
                         exception_handling, xx))
                con.commit()
                con.close()
            except sqlite3.OperationalError:
                QMessageBox.critical(self, "错误", "无写入数据权限，请以管理员身份运行！")

        def time_judgment(target_time):
            """判断时间是否大于当前时间"""
            # 获取当前时间年月日和时分秒
            print(target_time)
            now_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
            # 将now_time转换为时间格式
            now_time = datetime.datetime.strptime(now_time, '%Y/%m/%d %H:%M:%S')
            # 将字符参数转换为时间格式
            target_time = datetime.datetime.strptime(target_time, '%Y/%m/%d %H:%M:%S')
            # 判断是否重新输入
            print(now_time)
            print(target_time)
            xx = 0
            if now_time < target_time:
                print('目标时间大于当前时间，正确')
                xx = 0
            else:
                print('目标时间小于当前时间，错误')
                xx = 1
            return xx

        # 打印当前tab页
        print(self.tabWidget.currentIndex())
        # 判断当前tab页
        # 读取功能区的参数
        repeat_number = self.spinBox.value()
        exception_handling = self.comboBox_9.currentText()
        # 图像点击事件的参数获取
        if self.tabWidget.currentIndex() == 0:
            # 获取5个参数命令，写入数据库
            instruction = "图像点击"
            image = self.comboBox_8.currentText() + '/' + self.comboBox.currentText()
            print(image)
            parameter_1 = self.comboBox_2.currentText()
            # 如果复选框被选中，则获取第二个参数
            if self.checkBox_2.isChecked():
                parameter_2 = '自动略过'
            else:
                parameter_2 = None
            # 将命令写入数据库
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            image=image, parameter_1=parameter_1,
                                            parameter_2=parameter_2)
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
                                            parameter_2=parameter_2)

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
                                            parameter_2=parameter_2)
            print('已经保存鼠标移动的数据至数据库')
        # 等待事件的参数获取
        elif self.tabWidget.currentIndex() == 3:
            instruction = "等待"
            # 获取等待的参数
            # 如果checkBox没有被选中，则第一个参数为等待时间
            parameter_1 = None
            parameter_2 = None
            if not self.checkBox.isChecked():
                parameter_1 = "等待"
                try:
                    parameter_2 = int(self.lineEdit_2.text())
                except ValueError:
                    QMessageBox.critical(self, "错误", "等待时间请输入数字！")
                    return
            elif self.checkBox.isChecked():
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

            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2)
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
                                            parameter_2=parameter_2)
            print('已经保存鼠标滚轮滑动的数据至数据库')
        # 文本输入事件的参数获取
        elif self.tabWidget.currentIndex() == 5:
            # 获取5个参数命令
            instruction = "文本输入"
            # 获取文本输入的参数
            # 文本输入的内容
            parameter_1 = self.textEdit.toPlainText()
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1)
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
                                            parameter_1=parameter_1)
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
                                            parameter_2=parameter_2)
        # 鼠标当前位置事件的参数获取
        elif self.tabWidget.currentIndex() == 8:
            instruction = "鼠标事件"
            # 获取鼠标当前位置的参数
            parameter_1 = self.comboBox_7.currentText()
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1)
        # excel信息录入功能的参数获取
        elif self.tabWidget.currentIndex() == 9:
            instruction = "excel信息录入"
            # 获取excel工作簿路径和工作表名称
            parameter_1 = self.comboBox_12.currentText() + "-" + self.comboBox_13.currentText()
            # 获取图像文件路径
            image = self.comboBox_14.currentText() + '/' + self.comboBox_15.currentText()
            # 获取单元格值
            parameter_2 = self.lineEdit_4.text()
            # 判断是否递增行号
            parameter_3 = self.checkBox_3.isChecked()
            writes_commands_to_the_database(instruction=instruction,
                                            repeat_number=repeat_number,
                                            exception_handling=exception_handling,
                                            parameter_1=parameter_1,
                                            parameter_2=parameter_2,
                                            parameter_3=parameter_3,
                                            image=image)

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
        # 添加分支表名
        self.pushButton_7.clicked.connect(lambda: self.select_file("分支表名"))
        # 添加扩展程序
        self.pushButton_9.clicked.connect(lambda: self.select_file("扩展程序"))
        # 删除listview中的项
        self.pushButton_2.clicked.connect(lambda: self.delete_listview(self.listView, "图像文件夹路径"))
        self.pushButton_4.clicked.connect(lambda: self.delete_listview(self.listView_2, "工作簿路径"))
        self.pushButton_8.clicked.connect(lambda: self.delete_listview(self.listView_4, "分支表名"))
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
        elif judge == "分支表名":
            # 弹出对话框输入分支表名
            text, ok = QInputDialog.getText(self, "输入分支表名", "请输入分支表名：")
            if ok:
                self.write_to_database(None, None, text, None)
        elif judge == "扩展程序":
            # 弹出对话框输入扩展程序
            text, ok = QInputDialog.getText(self, "输入扩展程序", "请输入扩展程序：")
            if ok:
                self.write_to_database(None, None, None, text)
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
        elif judge == '分支表名':
            c.execute("DELETE FROM 全局参数 WHERE 分支表名 = ?", (value,))
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

        def add_listview(list_, listview):
            """添加listview"""
            listview.setModel(QStandardItemModel())
            if len(list_) != 0:
                model_2 = listview.model()
                for i in list_:
                    model_2.appendRow(QStandardItem(i))

        add_listview(image_folder_path, self.listView)
        add_listview(excel_folder_path, self.listView_2)
        add_listview(branch_table_name, self.listView_4)
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
        print("全局参数读取成功！")
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
