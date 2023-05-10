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
import ctypes, sys
import datetime
import json
import os
import sqlite3
import sys
import time
import webbrowser

import cryptocode
import keyboard
import pyautogui
import requests
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, \
    QFileDialog, QTableWidgetItem, QMessageBox, QHeaderView, QDialog
from pyscreeze import unicode

# from main_work import mainWork, exit_main_work
from main_work_2 import MainWork, exit_main_work
from 窗体.about import Ui_Dialog
# from 窗体.add_instruction import Ui_Form
from 窗体.mainwindow import Ui_MainWindow
from 窗体.navigation import Ui_navigation
from 窗体.setting import Ui_Setting
from 窗体.info import Ui_Form

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
        self.main_work = MainWork(fil_path, self)
        # 实例化导航页窗口
        self.navigation = Na()
        # self.dialog_na = Na()
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
        # 获取数据，修改按钮
        self.toolButton_5.clicked.connect(self.get_data)
        # 获取数据，子窗体取消按钮
        self.navigation.pushButton_3.clicked.connect(self.get_data)
        # 获取数据，子窗体保存按钮
        self.navigation.pushButton_2.clicked.connect(self.get_data)
        # 删除数据，删除按钮
        self.pushButton_2.clicked.connect(self.delete_data)
        # 交换数据，上移按钮
        self.toolButton_3.clicked.connect(lambda: self.go_up_down("up"))
        self.toolButton_4.clicked.connect(lambda: self.go_up_down("down"))
        # 保存按钮
        self.actionb.triggered.connect(self.save_data_to_current)
        # 清空指令按钮
        self.toolButton_6.clicked.connect(self.clear_table)
        # 导入数据按钮
        self.actionf.triggered.connect(self.data_import)
        # 主窗体开始按钮
        self.pushButton_5.clicked.connect(self.start)
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

    def format_table(self):
        """设置主窗口表格格式"""
        list_tableview = [self.tableWidget, self.tableWidget_2, self.tableWidget_3, self.tableWidget_4,
                          self.tableWidget_5, self.tableWidget_6]
        # 列的大小拉伸，可被调整
        for i in list_tableview:
            # 显示标题
            i.horizontalHeader().show()
            i.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            # 列的大小为可交互式的，用户可以调整
            i.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
            # 列的大小调整为固定，列宽不会改变
            i.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
            i.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
            # 设置列宽为50像素
            i.setColumnWidth(6, 50)
            i.setColumnWidth(4, 70)
            i.setColumnWidth(0, 100)

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
        print("导航页窗口开启")
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
        cursor.execute('select 图像名称,键鼠命令,参数,参数2,重复次数,ID from 命令')
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
        xx = self.tableWidget.item(row, 5).text()
        print(row, column, xx)
        # 删除数据库中指定id的数据
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('delete from 命令 where ID=?', (xx,))
        con.commit()
        con.close()
        # 调用get_data()函数，刷新表格
        self.get_data()

    def go_up_down(self, judge):
        """向上或向下移动选中的行"""
        # 获取选中值的行号和id
        row = self.tableWidget.currentRow()
        column = self.tableWidget.currentColumn()
        try:
            xx = self.tableWidget.item(row, 5).text()
            print(row, column, xx)
            # 将选中行的数据在数据库中与上一行数据交换，如果是第一行则不交换
            id = int(self.tableWidget.item(row, 5).text())
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
                cursor.execute('select 图像名称,键鼠命令,参数,参数2,重复次数 from 命令 where ID=?', (id,))
                list_id = cursor.fetchall()
                cursor.execute('select 图像名称,键鼠命令,参数,参数2,重复次数 from 命令 where ID=?', (id_up_down,))
                list_id_up = cursor.fetchall()
                # 交换选中行和上一行的数据
                cursor.execute('update 命令 set 图像名称=?,键鼠命令=?,参数=?,参数2=?,重复次数=? where ID=?', (
                    list_id_up[0][0], list_id_up[0][1], list_id_up[0][2], list_id_up[0][3], list_id_up[0][4], id))
                cursor.execute('update 命令 set 图像名称=?,键鼠命令=?,参数=?,参数2=?,重复次数=? where ID=?',
                               (list_id[0][0], list_id[0][1], list_id[0][2], list_id[0][3], list_id[0][4], id_up_down))
                con.commit()
                con.close()
                execute_sql = False
            # 调用get_data()函数，刷新表格
            self.get_data()
            # 将焦点移动到交换后的行
            self.tableWidget.setCurrentCell(row_up_down, column)
        except AttributeError:
            pass

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
        self.main_work.star_work()
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

    def open_readme(self):
        """打开使用说明"""
        os.popen('README.pdf')


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

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)
        # 是否激活自定义点击次数
        # 如果comboBox_3中的文本为“左键（自定义次数）”，则激活spinBox_2,否则禁用
        self.comboBox_3.currentTextChanged.connect(self.spinBox_2_enable)

        # 添加文件夹选择按钮事件
        # self.na_Path = ''
        self.pushButton.clicked.connect(lambda: self.select_file(0))
        # 添加保存按钮事件
        self.pushButton_2.clicked.connect(self.save_data)
        # 获取鼠标位置参数
        self.pushButton_4.clicked.connect(self.mouseMoveEvent)
        # 设置当前日期和时间
        self.pushButton_5.clicked.connect(self.get_now_date_time)
        # 当按钮按下时，获取按键的名称
        self.pushButton_6.clicked.connect(self.print_key_name)

    def select_file(self, judge):
        """选择文件夹并返回文件夹名称"""
        global fil_path
        if judge == 0:
            # self.na_Path = QFileDialog.getExistingDirectory(self, "选择存储目标图像的文件夹")
            fil_path = QFileDialog.getExistingDirectory(self, "选择存储目标图像的文件夹")
        try:
            images_name = os.listdir(fil_path)
        except FileNotFoundError:
            images_name = []
        # 去除文件夹中非png文件名称
        for i in range(len(images_name) - 1, -1, -1):
            if ".png" not in images_name[i]:
                images_name.remove(images_name[i])
        print(images_name)
        print(fil_path)
        # self.label_3.setText(fil_path.split('/')[-1])
        self.label_3.setText(fil_path)
        self.comboBox.addItems(images_name)

    def print_key_name(self):
        pressed_keys = set()  # create an empty set to store pressed keys
        # 禁用当前按钮
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
            # 激活当前按钮
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

    def save_data(self):
        """获取4个参数命令，并保存至数据库"""

        def writes_commands_to_the_database(instruction, image, parameter, parameter_2, repeat_number):
            """向数据库写入命令"""
            # if fil_path != '':
            con = sqlite3.connect('命令集.db')
            cursor = con.cursor()
            try:
                cursor.execute('INSERT INTO 命令(图像名称,键鼠命令,参数,参数2,重复次数) VALUES (?,?,?,?,?)',
                               (image, instruction, parameter, parameter_2, repeat_number))
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
        # 判断当前tab页，并获取5个参数命令，保存至数据库
        # 图像点击事件的参数获取
        if self.tabWidget.currentIndex() == 0:
            # 获取5个参数命令
            instruction = "图像点击"
            image = self.comboBox.currentText()
            repeat_number = self.spinBox.value()
            parameter = self.comboBox_2.currentText()
            # 如果复选框被选中，则获取第二个参数
            if self.checkBox_2.isChecked():
                parameter_2 = '自动略过'
            else:
                parameter_2 = ''
            writes_commands_to_the_database(instruction, image, parameter, parameter_2, repeat_number)
            print('已经保存图像识别点击的数据至数据库')
        # 鼠标点击事件的参数获取
        elif self.tabWidget.currentIndex() == 1:
            instruction = "坐标点击"
            images = 0
            repeat_number = self.spinBox.value()
            parameter = self.comboBox_3.currentText()
            parameter_2 = self.label_9.text() + "-" + self.label_10.text() + "-" + str(self.spinBox_2.value())
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)

        elif self.tabWidget.currentIndex() == 2:
            # 获取5个参数命令
            instruction = "鼠标移动"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取鼠标移动的参数
            # 鼠标移动的方向
            parameter = self.comboBox_4.currentText()
            # 鼠标移动的距离
            try:
                parameter_2 = int(self.lineEdit.text())
            except ValueError:
                QMessageBox.critical(self, "错误", "移动距离请输入数字！")
                return
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)
            print('已经保存鼠标移动的数据至数据库')
        # 等待事件的参数获取
        elif self.tabWidget.currentIndex() == 3:
            instruction = "等待"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取等待的参数
            # 如果checkBox没有被选中，则第一个参数为等待时间
            parameter = ''
            parameter_2 = ''
            if not self.checkBox.isChecked():
                parameter = "等待"
                try:
                    parameter_2 = int(self.lineEdit_2.text())
                except ValueError:
                    QMessageBox.critical(self, "错误", "等待时间请输入数字！")
                    return
            elif self.checkBox.isChecked():
                parameter = "等待到指定时间"
                # 判断时间是否大于当前时间
                parameter_2 = self.dateTimeEdit.text() + "+" + self.comboBox_6.currentText()
                try:
                    xx = time_judgment(parameter_2.split('+')[0])
                    if xx == 1:
                        raise TimeoutError("Invalid number!")
                except TimeoutError:
                    QMessageBox.critical(self, "错误", "启动时间小于当前系统时间，无效的指令。")
                    return

            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)
        # 鼠标滚轮滑动事件的参数获取
        elif self.tabWidget.currentIndex() == 4:
            # 获取5个参数命令
            instruction = "滚轮滑动"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取鼠标滚轮滑动的参数
            # 鼠标滚轮滑动的方向
            parameter = self.comboBox_5.currentText()
            # 鼠标滚轮滑动的距离
            try:
                parameter_2 = self.lineEdit_3.text()
            except ValueError:
                QMessageBox.critical(self, "错误", "滑动距离请输入数字！")
                return
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)
            print('已经保存鼠标滚轮滑动的数据至数据库')
        # 文本输入事件的参数获取
        elif self.tabWidget.currentIndex() == 5:
            # 获取5个参数命令
            instruction = "文本输入"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取文本输入的参数
            # 文本输入的内容
            parameter = self.textEdit.toPlainText()
            parameter_2 = ''
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)
            print('已经保存文本输入的数据至数据库')
        # 按下键盘事件的参数获取
        elif self.tabWidget.currentIndex() == 6:
            instruction = "按下键盘"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取按下键盘的参数
            # 按下键盘的内容
            parameter = self.label_31.text()
            parameter_2 = ''
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)
            print('已经保存按键的数据至数据库')
        # 中键激活事件的参数获取
        elif self.tabWidget.currentIndex() == 7:
            instruction = "中键激活"
            images = 0
            repeat_number = self.spinBox.value()
            # 获取中键激活的参数
            # 中键激活的内容
            parameter = ''
            parameter_2 = ''
            if self.radioButton.isChecked():
                parameter = '模拟点击'
                parameter_2 = self.spinBox_3.value()
            elif self.radioButton_2.isChecked():
                parameter = '自定义'
                parameter_2 = ''
            writes_commands_to_the_database(instruction, images, parameter, parameter_2, repeat_number)

        # 关闭窗体
        self.close()


class Info(QDialog, Ui_Form):
    def __init__(self, parent=None):
        super(Info, self).__init__(parent)
        self.setupUi(self)
        self.setWindowModality(Qt.ApplicationModal)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)


if __name__ == "__main__":
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
