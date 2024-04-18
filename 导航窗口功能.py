import datetime
import os
import random
import re
import sqlite3

import openpyxl
import pyautogui
from PyQt5.QtCore import QUrl, QRegExp, Qt
from PyQt5.QtGui import QDesktopServices, QImage, QPixmap, QIntValidator, QRegExpValidator
from PyQt5.QtWidgets import QMessageBox, QButtonGroup, QTreeWidgetItemIterator, QFileDialog, QWidget, QApplication
from dateutil.parser import parse
from openpyxl.utils.exceptions import InvalidFileException
from pygments import highlight
from pygments.formatters import HtmlFormatter
from pygments.lexers import PythonLexer

from 功能类 import OutputMessage, TransparentWindow, ImageClick, CoordinateClick, PlayVoice, WaitWindow, \
    DialogWindow, WindowControl, GetTimeValue, GetExcelCellValue, RunPython, RunExternalFile, TextRecognition, \
    VerificationCode, SendWeChat
from 变量池窗口 import VariablePool_Win
from 截图模块 import ScreenCapture
from 数据库操作 import extract_global_parameter, extract_excel_from_global_parameter, get_branch_count, \
    sqlitedb, close_database, set_window_size, save_window_size, get_variable_info, get_ocr_info
from 窗体.导航窗口 import Ui_navigation
from 网页操作 import WebOption
from 设置窗口 import Setting
from 选择窗体 import Branch_exe_win


class Na(QWidget, Ui_navigation):
    """导航页窗体及其功能"""

    def __init__(self, main_window_=None):
        super().__init__(main_window_)
        self.main_window = main_window_
        self.out_mes = OutputMessage(None, self)  # 输出信息
        self.setupUi(self)
        # 去除最大化最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)
        set_window_size(self)  # 获取上次退出时的窗口大小
        self.tabWidget.setCurrentIndex(0)  # 设置默认页
        self.treeWidget.expandAll()  # treeWidget全部展开
        self.variable_sel_win = Branch_exe_win(self, '变量选择')  # 变量选择窗口
        self.lineEdit_22.textChanged.connect(self.on_find_item)  # 指令搜索功能
        self.transparent_window = TransparentWindow()  # 框选窗口
        # 添加保存按钮事件
        self.modify_id = None
        self.modify_row = None
        self.image_path = None
        self.parameter_1 = None  # 用于存储参数
        self.pushButton_2.clicked.connect(lambda: self.save_data())
        # 获取鼠标位置参数
        self.mouse_position_function = None
        # 调整异常处理选项时，控制窗口控件的状态
        self.comboBox_9.activated.connect(lambda: self.exception_handling_judgment_type('报错处理'))
        self.comboBox_10.activated.connect(lambda: self.exception_handling_judgment_type('分支名称'))
        self.combo_image_preview = {  # 需要图像预览功能
            '图像点击': (self.comboBox_8, self.comboBox),
            '图像等待': (self.comboBox_17, self.comboBox_18),
            '信息录入': (self.comboBox_14, self.comboBox_15),
        }
        self.combo_excel_preview = {  # 需要加载excel表格的功能
            '信息录入': (self.comboBox_12, self.comboBox_13),
            '网页录入': (self.comboBox_20, self.comboBox_23),
            '获取Excel': (self.comboBox_45, self.comboBox_46),
            '写入单元格': (self.comboBox_57, self.comboBox_58),
        }
        self.variable_input_control = {  # 需要插入变量的控件
            '文本输入': self.textEdit,
            '元素控制': self.textEdit_3,
            '发送消息': self.textEdit_2,
            '提示音': self.textEdit_4,
            '运行Python': self.textEdit_5,
            '写入单元格': self.textEdit_6,
        }
        self.branch_jump_control = {  # 需要分支跳转的功能
            '功能区参数': (self.comboBox_10, self.comboBox_11),
            '跳转分支': (self.comboBox_37, self.comboBox_38),
            '变量判断': (self.comboBox_52, self.comboBox_53),
            '按键等待': (self.comboBox_41, self.comboBox_42),
        }
        self.pushButton_9.clicked.connect(lambda: self.on_button_clicked('查看'))
        self.pushButton_10.clicked.connect(lambda: self.on_button_clicked('删除'))
        # 快捷选择导航页
        self.tab_title_list = [self.tabWidget.tabText(x) for x in range(self.tabWidget.count())]
        self.treeWidget.itemClicked.connect(
            lambda: self.switch_navigation_page(self.treeWidget.currentItem().text(0))
        )
        self.tabWidget.currentChanged.connect(self.tab_widget_change)
        # 映射标签标题和对应的函数
        self.function_mapping = {
            '图像点击': (lambda x: self.image_click_function(x), True),
            '坐标点击': (lambda x: self.coordinate_click_function(x), False),
            '移动鼠标': (lambda x: self.move_mouse_function(x), False),
            '时间等待': (lambda x: self.time_waiting_function(x), False),
            '图像等待': (lambda x: self.image_waiting_function(x), True),
            '滚轮滑动': (lambda x: self.scroll_wheel_function(x), False),
            '文本输入': (lambda x: self.text_input_function(x), False),
            '按下键盘': (lambda x: self.press_keyboard_function(x), False),
            '中键激活': (lambda x: self.middle_activation_function(x), False),
            '鼠标点击': (lambda x: self.mouse_click_function(x), False),
            '鼠标拖拽': (lambda x: self.mouse_drag_function(x), False),
            '信息录入': (lambda x: self.information_entry_function(x), True),
            '打开网址': (lambda x: self.open_web_page_function(x), False),
            '元素控制': (lambda x: self.ele_control_function(x), True),
            '网页录入': (lambda x: self.web_entry_function(x), True),
            '切换frame': (lambda x: self.toggle_frame_function(x), False),
            '保存表格': (lambda x: self.save_form_function(x), True),
            '拖动元素': (lambda x: self.drag_element_function(x), True),
            '全屏截图': (lambda x: self.full_screen_capture_function(x), False),
            '切换窗口': (lambda x: self.switch_window_function(x), False),
            '发送消息': (lambda x: self.wechat_function(x), False),
            '数字验证码': (lambda x: self.verification_code_function(x), True),
            '提示音': (lambda x: self.play_voice_function(x), False),
            '倒计时窗口': (lambda x: self.wait_window_function(x), False),
            '提示窗口': (lambda x: self.dialog_window_function(x), False),
            '跳转分支': (lambda x: self.branch_jump_function(x), False),
            '终止流程': (lambda x: self.termination_process_function(x), False),
            '窗口控制': (lambda x: self.window_control_function(x), True),
            '按键等待': (lambda x: self.key_wait_function(x), False),
            '获取时间': (lambda x: self.gain_time_function(x), False),
            '获取Excel': (lambda x: self.gain_excel_function(x), False),
            '获取对话框': (lambda x: self.get_dialog_function(x), False),
            '变量判断': (lambda x: self.contrast_variables_function(x), False),
            '运行Python': (lambda x: self.run_python_function(x), False),
            '运行外部文件': (lambda x: self.run_external_file_function(x), True),
            '写入单元格': (lambda x: self.input_cell_function(x), True),
            'OCR识别': (lambda x: self.ocr_recognition_function(x), False),
            '获取鼠标位置': (lambda x: self.get_mouse_position_function(x), False),
        }
        # 加载功能窗口的按钮功能
        for func_name in self.function_mapping:
            self.function_mapping[func_name][0]('按钮功能')
        # 加载第一个功能窗口的控件信息
        self.function_mapping[self.tabWidget.tabText(0)][0]('加载信息')
        self.tabWidget_2.setCurrentIndex(0)  # 设置到功能页面到预览页
        # 设置窗口的flag，否则加载异常
        flags = self.windowFlags()
        self.setWindowFlags(flags | Qt.Window)

    def showEvent(self, a0) -> None:
        """显示窗口时，加载功能窗口的主要功能"""
        self.lineEdit_5.clear()  # 清空备注
        self.textBrowser.clear()  # 清空测试输出
        self.comboBox_9.setCurrentIndex(0)  # 异常处理方式
        self.comboBox_10.clear()  # 分支表名

    def closeEvent(self, a0) -> None:
        """关闭窗口时,触发的动作"""
        if self.transparent_window.isVisible():  # 关闭框选窗口
            self.transparent_window.close()
        self.main_window.get_data(self.modify_row)
        # 窗口大小
        save_window_size((self.width(), self.height()), self.windowTitle())

    def on_find_item(self, filter_txt):
        """指令搜索功能"""
        it = QTreeWidgetItemIterator(self.treeWidget)
        while it.value():
            if filter_txt in it.value().text(0):
                it.value().setHidden(False)
                item = it.value()
                while item.parent():
                    item.parent().setHidden(False)
                    item = item.parent()
            else:
                it.value().setHidden(True)
            it += 1

    @staticmethod
    def select_groupBox(selected_groupBox, all_groupBoxes):
        """选择groupBox，当tab命令页中的groupBox被选中时，其他groupBox不被选中"""
        for groupBox in all_groupBoxes:
            groupBox.setChecked(groupBox == selected_groupBox)

    def switch_navigation_page(self, name, restore_parameters=None):
        """弹出窗口自动选择对应功能页
        :param name: 功能页名称
        :param restore_parameters: 恢复参数，元组：(图像路径，参数，重复次数，异常处理，备注)"""

        def reverse_exception_handling_judgment(exception_handling_text):
            """将异常处理方式还原到窗体控件"""
            # 判断异常处理方式
            if exception_handling_text in {'自动跳过', '提示异常并暂停', '提示异常并停止'}:
                self.comboBox_9.setCurrentText(exception_handling_text)
            elif '-' in exception_handling_text:
                # 处理跳转分支的情况
                select_branch_table_name, branch_index = exception_handling_text.split('-')
                self.comboBox_9.setCurrentText('跳转分支')
                # 解除异常处理方式的禁用，加载分支表名
                self.comboBox_10.addItems(extract_global_parameter('分支表名'))
                self.find_controls('分支', '功能区参数')
                self.comboBox_10.setEnabled(True)
                self.comboBox_11.setEnabled(True)
                # 设置分支表名和分支序号
                self.comboBox_10.setCurrentText(select_branch_table_name)
                self.comboBox_11.setCurrentText(branch_index)

        try:
            tab_index = self.tab_title_list.index(name)
            self.tabWidget.setCurrentIndex(tab_index)
            if restore_parameters:  # 如果有恢复参数
                self.lineEdit_5.setText(restore_parameters[4])
                self.spinBox.setValue(int(restore_parameters[2]))
                # 恢复参数
                self.image_path = restore_parameters[0]
                self.parameter_1 = eval(restore_parameters[1])
                func_selected = self.function_mapping.get(name)[0]  # 获取当前页的功能
                try:
                    reverse_exception_handling_judgment(restore_parameters[3])  # 恢复异常处理参数
                    func_selected('还原参数')
                    self.show_image_to_label(name)  # 显示图像
                except Exception as e:
                    print('还原参数错误', e)
        except ValueError:  # 如果没有找到对应的功能页，则跳过
            pass

    def get_test_dic(self,
                     repeat_number_,
                     image_=None,
                     parameter_1_=None,
                     parameter_2_=None,
                     parameter_3_=None,
                     parameter_4_=None
                     ):
        """返回测试字典,用于测试按钮的功能"""
        self.tabWidget_2.setCurrentIndex(2)
        return {
            'ID': None,
            '图像路径': image_,
            '参数1（键鼠指令）': str(parameter_1_),
            '参数2': parameter_2_,
            '参数3': parameter_3_,
            '参数4': parameter_4_,
            '重复次数': repeat_number_
        }

    def get_func_info(self) -> dict:
        """返回功能区的参数"""

        def exception_handling_judgment():
            """判断异常处理方式
            :return: 异常处理方式"""
            exception_handling_text = None
            selected_text = self.comboBox_9.currentText()
            if selected_text in {'自动跳过', '提示异常并暂停', '提示异常并停止'}:
                exception_handling_text = selected_text
            elif selected_text == '跳转分支':
                select_branch_table_name = self.comboBox_10.currentText()
                if self.comboBox_11.currentText() == '':
                    QMessageBox.critical(self, "错误", "分支表下无指令，请检查分支表名是否正确！")
                    raise ValueError
                exception_handling_text = f'{select_branch_table_name}-{int(self.comboBox_11.currentText())}'
            return exception_handling_text

        # 当前页的index
        tab_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        return {
            '重复次数': self.spinBox.value(),
            '异常处理': exception_handling_judgment(),
            '备注': self.lineEdit_5.text(),
            '指令类型': tab_title
        }

    def find_controls(self, type_, ins_name: str) -> None:
        """加载不同的控件变量
        :param type_: 加载类型（图像、excel、分支）
        :param ins_name: 指令的名称"""

        def find_images() -> None:
            """选择图像文件夹并返回文件夹名称"""
            combox_folder, combox_file_name = self.combo_image_preview.get(ins_name)
            folder_path = combox_folder.currentText()
            try:
                # List all files in folder_path
                images_name = [f for f in os.listdir(folder_path) if f.endswith('.png')]
                # Sort files by modification time
                images_name.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)
            except FileNotFoundError:
                images_name = []
            # 清空combox_2中的所有元素
            combox_file_name.clear()
            # 将images_name中的所有元素添加到combox_2中
            combox_file_name.addItems(images_name)
            self.label_3.setText(self.comboBox_8.currentText())
            QApplication.processEvents()

        def find_excel_sheet_name() -> None:
            """获取excel表格中的所有sheet名称"""
            comboBox_before, comboBox_after = self.combo_excel_preview.get(ins_name)
            excel_path = comboBox_before.currentText()
            try:
                # 用openpyxl获取excel表格中的所有sheet名称
                excel_sheet_name = openpyxl.load_workbook(excel_path).sheetnames
            except FileNotFoundError:
                excel_sheet_name = []
            except InvalidFileException:
                excel_sheet_name = []
            except PermissionError:
                QMessageBox.critical(self, "错误", "当前文件被占用，请关闭文件后重试！")
                excel_sheet_name = []
            comboBox_after.clear()
            comboBox_after.addItems(excel_sheet_name)

        def find_branch_count() -> None:
            """当分支表名改变时，加载分支中的命令序号"""
            comboBox_branch_name, comboBox_branch_order = self.branch_jump_control.get(ins_name)
            count_record_ = get_branch_count(comboBox_branch_name.currentText())
            comboBox_branch_order.clear()
            # 加载分支中的命令序号
            branch_order_ = [str(i) for i in range(1, count_record_ + 1)]
            if len(branch_order_) != 0:
                comboBox_branch_order.addItems(branch_order_)

        if type_ == '图像':
            find_images()
        elif type_ == 'excel':
            find_excel_sheet_name()
        elif type_ == '分支':
            find_branch_count()

    def mouseMoveEvent(self, event):
        self.merge_additional_functions('get_mouse_position')

    def tab_widget_change(self):
        """切换导航页功能"""

        def control_status(disable_control_):
            """控制控件的状态，功能区参数控件的状态"""
            self.label_33.setVisible(disable_control_)
            self.label_34.setVisible(disable_control_)
            self.label_35.setVisible(disable_control_)
            self.comboBox_9.setCurrentIndex(0)
            self.comboBox_9.setVisible(disable_control_)
            self.comboBox_10.clear()
            self.comboBox_10.setVisible(disable_control_)
            self.comboBox_11.clear()
            self.comboBox_11.setVisible(disable_control_)

        try:
            # 获取当前活动页面的标题
            current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
            disable_control = self.function_mapping.get(current_title)[1]
            control_status(disable_control)  # 控制控件的状态
            if self.transparent_window.isVisible():  # 关闭框选窗口
                self.transparent_window.close()
                # 加载功能窗口的按钮功能
            self.function_mapping[current_title][0]('加载信息')
            self.tabWidget_2.setCurrentIndex(0)
        except TypeError:
            pass

    def merge_additional_functions(self, function_name, pars_1=None):
        """将一次性和冗余的功能合并
        :param pars_1:参数1
        :param function_name: 功能名称
        """
        if function_name == 'get_mouse_position':
            # 获取鼠标位置
            x, y = pyautogui.position()
            if self.mouse_position_function == '坐标点击':
                self.label_9.setText(str(x))
                self.label_10.setText(str(y))
            elif self.mouse_position_function == '开始拖拽':
                self.label_59.setText(str(x))
                self.label_61.setText(str(y))
            elif self.mouse_position_function == '结束拖拽':
                self.label_65.setText(str(x))
                self.label_66.setText(str(y))
        elif function_name == 'change_get_mouse_position_function':
            # 改变获取鼠标位置功能
            if pars_1 == '开始拖拽':
                self.mouse_position_function = '开始拖拽'
            elif pars_1 == '结束拖拽':
                self.mouse_position_function = '结束拖拽'
            elif pars_1 == '坐标点击':
                self.mouse_position_function = '坐标点击'
        elif function_name == '打开变量池':
            variable_pool = VariablePool_Win(self)
            variable_pool.exec_()
        elif function_name == '打开变量选择':
            self.variable_sel_win.show()

    def exception_handling_judgment_type(self, type_):
        """判断异常护理选项并调整控件
        :param type_: 判断类型（报错处理、分支名称）"""

        def disable_combobox(judge: bool = False):
            """禁用控件"""
            self.comboBox_10.clear()
            self.comboBox_10.setEnabled(judge)
            self.comboBox_11.clear()
            self.comboBox_11.setEnabled(judge)

        try:
            if type_ == '报错处理':  # 报错处理下拉列表变化触发
                if self.comboBox_9.currentText() == '自动跳过':
                    disable_combobox()
                elif self.comboBox_9.currentText() == '提示异常并暂停':
                    disable_combobox()
                elif self.comboBox_9.currentText() == '提示异常并停止':
                    disable_combobox()
                elif self.comboBox_9.currentText() == '跳转分支':
                    disable_combobox(True)
                    self.comboBox_10.addItems(extract_global_parameter('分支表名'))
                    self.comboBox_10.setCurrentIndex(0)
                    self.find_controls('分支', '功能区参数')
            elif type_ == '分支名称':  # 分支表名下拉列表变化触发
                self.find_controls('分支', '功能区参数')
        except sqlite3.OperationalError:
            pass

    def quick_screenshot(self, control_name, judge):
        """截图功能
        :param control_name: 需要的控件
        :param judge: 功能选择（快捷截图、打开文件夹）"""
        if judge == '快捷截图':
            if control_name.currentText() == '':
                QMessageBox.warning(self, '警告', '未选择图像文件夹！', QMessageBox.Yes)
            else:
                # 隐藏主窗口
                self.hide()
                self.main_window.hide()
                # 截图
                screen_capture = ScreenCapture()
                screen_capture.screenshot_area()  # 设置截图区域
                screen_capture.screenshot_region()  # 截图
                screen_capture.show_preview()  # 显示预览
                # 显示主窗口
                self.show()
                self.main_window.show()
                # 刷新图像文件夹
                QApplication.processEvents()
                current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
                self.find_controls('图像', current_title)
                self.show_image_to_label(current_title)

        elif judge == '打开文件夹':
            if control_name.currentText() != '':
                os.startfile(os.path.normpath(control_name.currentText()))

        elif judge == '设置区域':
            # 隐藏主窗口
            self.hide()
            self.main_window.hide()
            # 截图
            screen_capture = ScreenCapture()
            screen_capture.screenshot_area()  # 设置截图区域
            control_name.setText(str(screen_capture.region))
            # 显示主窗口
            self.show()
            self.main_window.show()

    def writes_commands_to_the_database(self,
                                        instruction_,
                                        repeat_number_,
                                        exception_handling_,
                                        image_=None,
                                        parameter_1_=None,
                                        parameter_2_=None,
                                        parameter_3_=None,
                                        parameter_4_=None,
                                        remarks_=None
                                        ):
        """向数据库写入命令"""
        try:
            cursor, con = sqlitedb()
            branch_name = self.main_window.comboBox.currentText()

            query_params = (
                image_, instruction_, str(parameter_1_), parameter_2_, parameter_3_, parameter_4_, repeat_number_,
                exception_handling_, remarks_, branch_name
            )
            if self.pushButton_2.text() == '添加指令':
                cursor.execute(
                    'INSERT INTO 命令'
                    '(图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) '
                    'VALUES (?,?,?,?,?,?,?,?,?,?)',
                    query_params
                )

            elif self.pushButton_2.text() == '修改指令':
                cursor.execute(
                    'UPDATE 命令 '
                    'SET 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=?,备注=?,隶属分支=? '
                    'WHERE ID=?',
                    query_params + (self.modify_id,)
                )

            elif self.pushButton_2.text() == '向前插入':
                # 将当前ID和之后的ID递增1
                max_id_ = 1000000
                cursor.execute('UPDATE 命令 SET ID=ID+? WHERE ID>=?', (max_id_, self.modify_id))
                cursor.execute('UPDATE 命令 SET ID=ID-? WHERE ID>=?', (max_id_ - 1, max_id_ + int(self.modify_id)))
                # 插入新的命令
                cursor.execute(
                    'INSERT INTO 命令'
                    '(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) '
                    'VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                    (self.modify_id,) + query_params
                )

            elif self.pushButton_2.text() == '向后插入':
                self.modify_row = self.modify_row + 1
                try:
                    cursor.execute(
                        'INSERT INTO 命令'
                        '(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) '
                        'VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                        (self.modify_id + 1,) + query_params
                    )
                except sqlite3.IntegrityError:
                    # 如果下一个id已经存在，则将后面的id全部加1
                    max_id_ = 1000000
                    cursor.execute('UPDATE 命令 SET ID=ID+? WHERE ID>?', (max_id_, self.modify_id))
                    cursor.execute('UPDATE 命令 SET ID=ID-? WHERE ID>?', (max_id_ - 1, max_id_ + int(self.modify_id)))
                    # 插入新的命令
                    cursor.execute(
                        'INSERT INTO 命令'
                        '(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) '
                        'VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                        (self.modify_id + 1,) + query_params
                    )

            con.commit()
            close_database(cursor, con)

        except sqlite3.OperationalError:
            QMessageBox.critical(self, "错误", "数据写入失败，请重试！")

    def save_data(self):
        """获取4个参数命令，并保存至数据库"""
        # 当前页的index
        tab_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        try:
            func_selected = self.function_mapping.get(tab_title)[0]  # 获取当前页的功能
            # 根据功能获取参数
            if func_selected:
                try:
                    func_selected('写入参数')
                    self.close()
                except Exception as e:
                    print(e)
        except TypeError:
            pass

    def show_image_to_label(self, ins_name: str, judge='显示'):
        """将图像显示到label中,图像预览的功能
        :param ins_name: 指令名称
        :param judge: 显示、删除、查看"""
        try:
            comboBox_folder, comboBox_image = self.combo_image_preview.get(ins_name)
            image_path = os.path.normpath(
                os.path.join(comboBox_folder.currentText(), comboBox_image.currentText())
            )
            if (os.path.exists(image_path)) and (os.path.isfile(image_path)):  # 判断图像是否存在
                if judge == '显示':
                    # 将图像转换为QImage对象
                    image_ = QImage(image_path)
                    image = image_.scaled(self.label_43.width(), self.label_43.height(), Qt.KeepAspectRatio)
                    self.label_43.setPixmap(QPixmap.fromImage(image))
                elif judge == '删除':
                    os.remove(image_path)
                elif judge == '查看':
                    os.startfile(image_path)
            else:
                self.label_43.setText('暂无')
            self.tabWidget_2.setCurrentIndex(1)  # 设置到功能页面到预览页
        except TypeError:
            pass

    def on_button_clicked(self, judge: str) -> None:
        """按钮点击事件,用于图像预览的按钮事件
        :param judge: 执行的操作(删除、查看)"""
        # 获取当前页的标题
        current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        if current_title in self.combo_image_preview:
            self.show_image_to_label(current_title, judge)
            if judge == '删除':
                self.find_controls('图像', current_title)
                self.show_image_to_label(current_title)

    def write_value_to_textedit(self, value: str) -> None:
        """将变量池中的值写入到文本框中"""

        def append_textedit(new_text):
            errorFormat_ = '<font color="red">{}</font>'
            # 使textEdit显示不同的文本
            current_title_ = self.tabWidget.tabText(self.tabWidget.currentIndex())
            textEdit = self.variable_input_control.get(current_title_)
            if textEdit.isEnabled():
                textEdit.insertHtml('☾')
                textEdit.insertHtml((errorFormat_.format(new_text)))
                textEdit.insertHtml('☽')

        if value:
            append_textedit(value)

    def image_click_function(self, type_):
        """图像点击识别窗口的功能
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = os.path.normpath(os.path.join(self.comboBox_8.currentText(), self.comboBox.currentText()))
            parameter_1_ = self.comboBox_2.currentText()
            # 如果复选框被选中，则获取第二个参数
            parameter_2_ = None
            if self.radioButton_2.isChecked():
                parameter_2_ = '自动略过'
            elif self.radioButton_4.isChecked():
                parameter_2_ = self.spinBox_4.value()
            if self.groupBox_57.isChecked():
                parameter_4_ = self.label_155.text()
            else:
                parameter_4_ = '(0,0,0,0)'
            # 从tab页获取参数
            parameter_dic_ = {
                '动作': parameter_1_,
                '异常': parameter_2_,
                '区域': parameter_4_,
                '灰度': self.checkBox.isChecked()
            }
            # 检查参数是否有异常
            if (os.path.isdir(image_)) or (not os.path.exists(image_)):
                QMessageBox.critical(self, "错误", "图像文件不存在，请检查图像文件是否存在！")
                raise FileNotFoundError
            if self.groupBox_57.isChecked() and self.label_155.text() == '(0,0,0,0)':
                QMessageBox.critical(self, "错误", "未设置识别区域！")
                raise FileNotFoundError

            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到窗体控件"""
            # 将图像路径设置回对应的comboBox
            image_folder, image_file = os.path.split(image_)
            image_folder_index = self.comboBox_8.findText(image_folder)
            if image_folder_index != -1:
                self.comboBox_8.setCurrentIndex(image_folder_index)
            else:
                # 如果路径不存在，则添加路径
                self.comboBox_8.addItem(image_folder)
                self.comboBox_8.setCurrentIndex(self.comboBox_8.findText(image_folder))

            image_file_index = self.comboBox.findText(image_file)
            if image_file_index != -1:
                self.comboBox.setCurrentIndex(image_file_index)
            else:
                # 如果文件不存在，则添加文件
                self.comboBox.addItem(image_file)
                self.comboBox.setCurrentIndex(self.comboBox.findText(image_file))

            # 将其他参数设置回对应的控件
            self.comboBox_2.setCurrentText(parameter_dic_['动作'])

            if parameter_dic_['异常'] == '自动略过':
                self.radioButton_2.setChecked(True)
            else:
                self.radioButton_4.setChecked(True)
                self.spinBox_4.setValue(parameter_dic_['异常'])

            if parameter_dic_['区域'] == '(0,0,0,0)':
                self.groupBox_57.setChecked(False)
            else:
                self.groupBox_57.setChecked(True)
                self.label_155.setText(parameter_dic_['区域'])

            self.checkBox.setChecked(parameter_dic_['灰度'])

        def test():
            """测试功能"""
            try:
                image_, parameter_1_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_,
                                         parameter_1_=parameter_1_
                                         )
                # 测试用例
                try:
                    image_click = ImageClick(self.out_mes, dic_)
                    image_click.is_test = True
                    image_click.start_execute()
                except Exception as e:
                    print(e)
                    self.out_mes.out_mes(f'未找到目标图像，测试结束', True)

            except FileNotFoundError:
                self.out_mes.out_mes(f'图像文件未设置！', True)

        def open_setting_window():
            """打开图像点击设置窗口"""
            setting_win = Setting(self)  # 设置窗体
            setting_win.setModal(True)
            setting_win.exec_()

        if type_ == '按钮功能':
            # 快捷截图功能
            self.pushButton.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_8, '快捷截图')
            )
            # 打开图像文件夹
            self.pushButton_7.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_8, '打开文件夹')
            )
            # 设置区域
            self.pushButton_50.clicked.connect(
                lambda: self.quick_screenshot(self.label_155, '设置区域')
            )
            # 加载下拉列表数据
            self.comboBox_8.activated.connect(
                lambda: self.find_controls('图像', '图像点击')
            )
            # 元素预览
            self.comboBox.activated.connect(
                lambda: self.show_image_to_label('图像点击')
            )
            # 测试按钮
            self.pushButton_6.clicked.connect(test)
            # 打开设置窗口
            self.pushButton_11.clicked.connect(open_setting_window)

        elif type_ == '写入参数':
            image, parameter_1 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_1,
                                                 remarks_=func_info_dic.get('备注'))

        elif type_ == '还原参数':
            put_parameters(self.image_path, self.parameter_1)

        elif type_ == '加载信息':
            # 加载图像文件夹路径
            self.comboBox_8.clear()
            self.comboBox_8.addItems(extract_global_parameter('资源文件夹路径'))
            self.find_controls('图像', '图像点击')
            self.show_image_to_label('图像点击')

    def scroll_wheel_function(self, type_):
        """滚轮滑动的窗口功能"""

        def put_parameters(parameter_dic__):
            """将参数还原到窗体控件"""
            parameter_1_ = self.parameter_1['类型']
            if parameter_1_ == '随机滚轮滑动':
                self.groupBox_29.setChecked(True)
                self.groupBox_22.setChecked(False)
                self.spinBox_16.setValue(parameter_dic__['最小距离'])
                self.spinBox_17.setValue(parameter_dic__['最大距离'])
            elif parameter_1_ == '滚轮滑动':
                self.groupBox_22.setChecked(True)
                self.groupBox_29.setChecked(False)
                self.comboBox_5.setCurrentText(parameter_dic__['方向'])
                self.lineEdit_3.setText(parameter_dic__['距离'])

        if type_ == '按钮功能':
            # 将不同的单选按钮添加到同一个按钮组
            all_groupBoxes_ = [self.groupBox_22, self.groupBox_29]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_))
            self.lineEdit_3.setValidator(QIntValidator())  # 设置只能输入数字

        elif type_ == '写入参数':
            parameter_dic_ = None
            parameter_1 = None
            if self.groupBox_22.isChecked():
                parameter_1 = '滚轮滑动'
            elif self.groupBox_29.isChecked():
                parameter_1 = '随机滚轮滑动'
            # 检查参数是否有异常
            if not self.lineEdit_3.text().isdigit() and self.groupBox_22.isChecked():
                QMessageBox.critical(self, "错误", "滚动的距离未输入！")
                raise ValueError
            # 参数字典
            if parameter_1 == '随机滚轮滑动':
                parameter_dic_ = {
                    '类型': parameter_1,
                    '最小距离': self.spinBox_16.value(),
                    '最大距离': self.spinBox_17.value()
                }
            elif parameter_1 == '滚轮滑动':
                parameter_dic_ = {
                    '类型': parameter_1,
                    '方向': self.comboBox_5.currentText(),
                    '距离': self.lineEdit_3.text()
                }
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 parameter_1_=parameter_dic_)
        elif type_ == '还原参数':
            put_parameters(self.parameter_1)

    def text_input_function(self, type_):
        """文本输入窗口的功能"""

        def check_text_type():
            # 检查文本输入类型
            text = self.textEdit.toPlainText()
            # 检查text中是否为英文大小写字母和数字
            if re.search('[a-zA-Z0-9]', text) is None:
                self.checkBox_2.setChecked(False)
                QMessageBox.warning(self, '警告', '特殊控件的文本输入仅支持输入英文大小写字母和数字！', QMessageBox.Yes)

        if type_ == '按钮功能':
            # 检查输入的数据是否合法
            self.checkBox_2.clicked.connect(check_text_type)
            self.pushButton_28.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))

        elif type_ == '写入参数':
            # 文本输入的内容
            parameter_1 = self.textEdit.toPlainText()
            parameter_2 = str(self.checkBox_2.isChecked())
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def coordinate_click_function(self, type_):
        """坐标点击识别窗口的功能
        :param type_: 功能名称（加载按钮、主要功能）"""

        def spinBox_2_enable():
            # 激活自定义点击次数
            if self.comboBox_3.currentText() == '左键（自定义次数）':
                self.spinBox_2.setEnabled(True)
                self.label_22.setEnabled(True)
            else:
                self.spinBox_2.setEnabled(False)
                self.label_22.setEnabled(False)
                self.spinBox_2.setValue(0)

        def test():
            """测试功能"""
            parameter_1_ = self.comboBox_3.currentText()
            parameter_2_ = f'{self.label_9.text()}-{self.label_10.text()}-{self.spinBox_2.value()}'
            dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                     parameter_1_=parameter_1_,
                                     parameter_2_=parameter_2_)
            # 测试用例
            try:
                cor_click = CoordinateClick(self.out_mes, dic_)
                cor_click.is_test = True
                cor_click.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'参数异常', True)

        if type_ == '按钮功能':
            # 坐标点击
            self.pushButton_4.pressed.connect(
                lambda: self.merge_additional_functions(
                    'change_get_mouse_position_function', '坐标点击'
                )
            )
            self.pushButton_4.clicked.connect(self.mouseMoveEvent)
            # 是否激活自定义点击次数
            self.comboBox_3.activated.connect(spinBox_2_enable)
            # 测试按钮
            self.pushButton_23.clicked.connect(test)
        elif type_ == '写入参数':
            parameter_1 = self.comboBox_3.currentText()
            parameter_2 = f'{self.label_9.text()}-{self.label_10.text()}-{self.spinBox_2.value()}'
            # 检查参数是否有异常
            if self.label_9.text() == '0' and self.label_10.text() == '0':
                QMessageBox.critical(self, "错误", "未设置坐标，请设置坐标！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def time_waiting_function(self, type_):
        """等待识别窗口的功能
        :param type_: 功能名称（加载按钮、主要功能）"""

        def time_judgment(target_time):
            """判断时间是否大于当前时间"""
            now_time = datetime.datetime.now()
            return True if now_time < parse(target_time) else False

        def get_now_date_time():
            """将当前的时间和日期设置为dateTimeEdit的日期和时间"""
            if self.radioButton_17.isChecked():
                # 获取当前日期和时间
                now_date_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                # 将当前的时间和日期加10分钟
                new_date_time = parse(now_date_time) + datetime.timedelta(minutes=10)
                # 将dateTimeEdit的日期和时间设置为当前日期和时间
                self.dateTimeEdit.setDateTime(new_date_time)

        if type_ == '按钮功能':
            # 将不同的单选按钮添加到同一个按钮组
            buttonGroup = QButtonGroup(self)
            buttonGroup.addButton(self.radioButton_16)
            buttonGroup.addButton(self.radioButton_18)
            buttonGroup.addButton(self.radioButton_17)
            # 设置当前日期和时间
            self.radioButton_17.clicked.connect(get_now_date_time)
        elif type_ == '写入参数':
            # 如果checkBox没有被选中，则第一个参数为等待时间
            parameter_1 = None
            parameter_2 = None
            # 时间等待
            if self.radioButton_16.isChecked():
                parameter_1 = "时间等待"
                parameter_2 = self.spinBox_13.value()
            # 等待到指定时间
            elif self.radioButton_17.isChecked():
                parameter_1 = "定时等待"
                # 判断时间是否大于当前时间
                parameter_2 = self.dateTimeEdit.text() + "+" + self.comboBox_6.currentText()
                if not time_judgment(self.dateTimeEdit.text()):
                    QMessageBox.critical(self, "错误", "定时时间小于当前系统时间，无效指令。")
                    raise ValueError
            # 随机等待
            elif self.radioButton_18.isChecked():
                parameter_1 = "随机等待"
                min_time = self.spinBox_14.value()
                max_time = self.spinBox_15.value()
                parameter_2 = str(min_time) + "-" + str(max_time)
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def image_waiting_function(self, type_):
        """图像等待识别窗口的功能"""
        if type_ == '按钮功能':
            # 下拉列表数据
            self.comboBox_17.activated.connect(
                lambda: self.find_controls('图像', '图像等待')
            )
            # 元素预览
            self.comboBox_18.activated.connect(
                lambda: self.show_image_to_label('图像等待')
            )
            # 快捷截图功能
            self.pushButton_21.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_17, '快捷截图')
            )
            # 打开图像文件夹
            self.pushButton_22.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_17, '打开文件夹')
            )
        elif type_ == '写入参数':
            # 获取参数
            image = os.path.normpath(os.path.join(self.comboBox_8.currentText(), self.comboBox.currentText()))
            parameter_1 = self.comboBox_19.currentText()
            parameter_2 = self.spinBox_6.value()
            # 检查参数是否有异常
            if (os.path.isdir(image)) or (not os.path.exists(image)):
                QMessageBox.critical(self, "错误", "图像文件不存在，请检查图像文件是否存在！")
                raise FileNotFoundError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 image_=image,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2)
        elif type_ == '加载信息':
            # 加载图像文件夹路径
            self.comboBox_17.clear()
            self.comboBox_18.clear()
            self.comboBox_17.addItems(extract_global_parameter('资源文件夹路径'))
            self.find_controls('图像', '图像等待')
            self.show_image_to_label('图像等待')

    def move_mouse_function(self, type_):
        """鼠标移动识别窗口的功能"""

        if type_ == '按钮功能':
            all_groupBoxes_ = [self.groupBox_28, self.groupBox_30, self.groupBox_59]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_))

            # 限制输入框只能输入数字
            self.lineEdit.setValidator(QIntValidator())
            self.lineEdit_29.setValidator(QIntValidator())
            self.lineEdit_30.setValidator(QIntValidator())

        elif type_ == '写入参数':
            parameter_1 = None
            parameter_2 = None
            # 鼠标移动
            if self.radioButton_19.isChecked():
                parameter_1 = '移动鼠标'
                parameter_2 = f'{self.comboBox_4.currentText()},{self.lineEdit.text()}'
            # 随机移动
            elif self.radioButton_ran.isChecked():
                parameter_1 = '随机移动鼠标'
                parameter_2 = self.comboBox_16.currentText()
            # 检查参数是否有异常
            if not self.lineEdit.text().isdigit() and self.radioButton_19.isChecked():
                QMessageBox.critical(self, "错误", "鼠标移动的距离未输入！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

        elif type_ == '加载信息':
            # 加载下拉列表数据
            self.comboBox_61.clear()
            self.comboBox_61.addItems(get_variable_info('list'))

    def press_keyboard_function(self, type_):
        """按下键盘识别窗口的功能"""

        if type_ == '按钮功能':
            # 当按钮按下时，获取按键的名称
            pass
        elif type_ == '写入参数':
            # 按下键盘的内容
            parameter_1 = self.keySequenceEdit.keySequence().toString()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 parameter_1_=parameter_1)

    def middle_activation_function(self, type_):
        """中键激活的窗口功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            # 中键激活的内容
            parameter_1 = None
            parameter_2 = None
            if self.radioButton.isChecked():
                parameter_1 = '模拟点击'
                parameter_2 = self.spinBox_3.value()
            elif self.radioButton_zi.isChecked():
                parameter_1 = '自定义'
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def mouse_click_function(self, type_):
        """鼠标点击的窗口的功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            # 获取鼠标当前位置的参数
            parameter_1 = self.comboBox_35.currentText().replace('（自定义次数）', '')
            parameter_2 = f'{self.spinBox_18.value()}-{self.spinBox_19.value()}-{self.spinBox_20.value()}'
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def information_entry_function(self, type_):
        """信息录入的窗口功能"""

        def line_number_increasing():
            # 行号递增功能被选中后弹出提示框
            if self.checkBox_3.isChecked():
                QMessageBox.information(self, '提示',
                                        '启用该功能后，请在主页面中设置循环次数大于1，执行全部指令后，'
                                        '循环执行时，单元格行号会自动递增。',
                                        QMessageBox.Ok
                                        )

        if type_ == '按钮功能':
            # 行号自动递增提示
            self.checkBox_3.clicked.connect(line_number_increasing)
            # 信息录入页面的快捷截图功能
            self.pushButton_5.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_14, '快捷截图')
            )
            self.pushButton_8.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_14, '打开文件夹')
            )
            # 信息录入窗口的excel功能
            self.comboBox_12.activated.connect(
                lambda: self.find_controls('excel', '信息录入')
            )
            # 加载下拉列表数据
            self.comboBox_14.activated.connect(
                lambda: self.find_controls('图像', '信息录入')
            )
            # 图像预览
            self.comboBox_15.activated.connect(
                lambda: self.show_image_to_label('信息录入')
            )
        elif type_ == '写入参数':
            parameter_4 = None
            # 获取excel工作簿路径和工作表名称
            parameter_1 = self.comboBox_12.currentText() + "-" + self.comboBox_13.currentText()
            # 获取图像文件路径
            image = os.path.normpath(self.comboBox_14.currentText() + '/' + self.comboBox_15.currentText())
            # 获取单元格值
            parameter_2 = self.lineEdit_4.text()
            # 判断是否递增行号和特殊控件输入
            parameter_3 = str(self.checkBox_3.isChecked()) + '-' + str(self.checkBox_4.isChecked())
            # 判断其他参数
            if self.radioButton_3.isChecked() and not self.radioButton_5.isChecked():
                parameter_4 = '自动跳过'
            elif not self.radioButton_3.isChecked() and self.radioButton_5.isChecked():
                parameter_4 = self.spinBox_5.value()
            # 检查参数是否有异常
            if (os.path.isdir(image)) or (not os.path.exists(image)):
                QMessageBox.critical(self, "错误", "图像文件不存在，请检查图像文件是否存在！")
                raise FileNotFoundError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 parameter_4_=parameter_4,
                                                 image_=image,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 加载文件路径
            self.comboBox_12.clear()
            self.comboBox_12.addItems(extract_excel_from_global_parameter())  # 加载全局参数中的excel文件路径
            self.comboBox_13.clear()
            self.comboBox_14.clear()
            self.comboBox_14.addItems(extract_global_parameter('资源文件夹路径'))
            self.find_controls('图像', '信息录入')
            self.show_image_to_label('信息录入')
            self.find_controls('excel', '信息录入')

    def open_web_page_function(self, type_):
        """打开网址的窗口功能"""

        def web_functional_testing(judge):
            """网页连接测试"""
            if judge == '测试':
                url = self.lineEdit_19.text()
                web_option = WebOption(self.out_mes)
                is_succeed, str_info = web_option.web_open_test(url)  # 测试网页是否能打开
                if is_succeed:
                    QMessageBox.information(self, '提示', '连接成功！', QMessageBox.Yes)
                else:
                    QMessageBox.critical(self, "错误", str_info)

            elif judge == '安装浏览器':
                url = 'https://google.cn/chrome/'
                QDesktopServices.openUrl(QUrl(url))

            elif judge == '安装浏览器驱动':
                # 弹出选择提示框
                x = QMessageBox.information(
                    self, '提示', '确认下载浏览器驱动？', QMessageBox.Yes | QMessageBox.No
                )
                if x == QMessageBox.Yes:
                    print('下载浏览器驱动')
                    self.tabWidget_2.setCurrentIndex(2)
                    web_option = WebOption(self.out_mes)
                    web_option.install_browser_driver()

        if type_ == '按钮功能':
            self.pushButton_18.clicked.connect(lambda: web_functional_testing('测试'))
            self.pushButton_19.clicked.connect(lambda: web_functional_testing('安装浏览器'))
            self.pushButton_20.clicked.connect(lambda: web_functional_testing('安装浏览器驱动'))
        elif type_ == '写入参数':
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 image_=self.lineEdit_19.text())

    def ele_control_function(self, type_):
        """网页元素控制的窗口功能"""

        def Lock_control():
            """锁定控件"""
            if self.comboBox_22.currentText() == '输入内容':
                self.textEdit_3.setEnabled(True)
            else:
                self.textEdit_3.clear()
                self.textEdit_3.setEnabled(False)

        if type_ == '按钮功能':
            Lock_control()
            self.comboBox_22.activated.connect(Lock_control)
            self.pushButton_31.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))

        elif type_ == '写入参数':
            element_type = self.comboBox_21.currentText()
            element_value = self.lineEdit_7.text()
            text = self.textEdit_3.toPlainText()
            action = self.comboBox_22.currentText()
            # 判断其他参数
            timeout_type = None
            if self.radioButton_6.isChecked() and not self.radioButton_7.isChecked():
                timeout_type = '自动跳过'
            elif not self.radioButton_6.isChecked() and self.radioButton_7.isChecked():
                timeout_type = self.spinBox_7.value()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 image_=element_type + '-' + element_value,
                                                 parameter_1_=action,
                                                 parameter_2_=text,
                                                 parameter_3_=timeout_type)

    def web_entry_function(self, type_):
        """网页录入的窗口功能"""
        if type_ == '按钮功能':
            # 网页信息录入的excel功能
            self.comboBox_20.activated.connect(
                lambda: self.find_controls('excel', '网页录入')
            )
        elif type_ == '写入参数':
            parameter_4 = None
            # 获取excel工作簿路径和工作表名称
            parameter_1 = self.comboBox_20.currentText() + "-" + self.comboBox_23.currentText()
            # 获取元素类型和元素
            image = self.comboBox_24.currentText().replace('：', '') + '-' + self.lineEdit_10.text()
            # 获取单元格值
            parameter_2 = self.lineEdit_9.text()
            # 判断是否递增行号和特殊控件输入
            parameter_3 = str(self.checkBox_6.isChecked())
            # 判断其他参数
            if self.radioButton_10.isChecked() and not self.radioButton_11.isChecked():
                parameter_4 = '自动跳过'
            elif not self.radioButton_10.isChecked() and self.radioButton_11.isChecked():
                parameter_4 = self.spinBox_8.value()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 parameter_4_=parameter_4,
                                                 image_=image,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 加载文件路径
            self.comboBox_20.clear()
            self.comboBox_20.addItems(extract_excel_from_global_parameter())
            self.comboBox_23.clear()
            self.find_controls('excel', '网页录入')

    def mouse_drag_function(self, type_):
        """鼠标拖拽窗口的功能"""

        def get_position(label_text, random_range=(-100, 100)):
            """获取鼠标位置"""
            if not self.checkBox_8.isChecked():
                position = int(label_text)
            else:
                position = int(label_text) + random.randint(*random_range)
            return position

        def get_random_offset(range_x, range_y=None):
            x_random = random.randint(*range_x)
            y_random = random.randint(*range_y) if range_y else 0
            return x_random, y_random

        def drag_test():
            # 开始拖拽，是否使用随机位置
            if not self.checkBox_8.isChecked():
                start_position = (int(self.label_59.text()), int(self.label_61.text()))
            else:
                x_offset, y_offset = get_random_offset((-100, 100))
                start_position = (int(self.label_59.text()) + x_offset, int(self.label_61.text()) + y_offset)
            # 结束拖拽，是否使用随机位置
            if not self.checkBox_7.isChecked():
                end_position = (int(self.label_65.text()), int(self.label_66.text()))
            else:
                x_offset, y_offset = get_random_offset((-200, 200), (-200, 200))
                end_position = (int(self.label_65.text()) + x_offset, int(self.label_66.text()) + y_offset)
            # 开始拖拽
            pyautogui.moveTo(start_position[0], start_position[1], duration=0.3)
            pyautogui.dragTo(end_position[0], end_position[1], duration=0.3)

        if type_ == '按钮功能':
            # 鼠标拖拽
            self.pushButton_12.pressed.connect(
                lambda: self.merge_additional_functions('change_get_mouse_position_function', '开始拖拽'))
            self.pushButton_12.clicked.connect(self.mouseMoveEvent)
            self.pushButton_13.pressed.connect(
                lambda: self.merge_additional_functions('change_get_mouse_position_function', '结束拖拽'))
            self.pushButton_13.clicked.connect(self.mouseMoveEvent)
            # 拖拽测试按钮
            self.pushButton_14.clicked.connect(drag_test)
        elif type_ == '写入参数':
            # 获取开始位置
            x_start = get_position(self.label_59.text())
            y_start = get_position(self.label_61.text())
            # 获取结束位置
            x_end = get_position(self.label_65.text(), random_range=(-200, 200))
            y_end = get_position(self.label_66.text(), random_range=(-200, 200))
            # 保存位置
            parameter_1 = f"{x_start},{y_start}"
            parameter_2 = f"{x_end},{y_end}"
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def toggle_frame_function(self, type_):
        """切换frame窗口的功能"""

        def switch_frame():
            """切换frame"""
            # 切换frame时控件的状态
            if self.comboBox_26.currentText() == '切换到指定frame':
                self.comboBox_27.setEnabled(True)
                self.lineEdit_11.clear()
                self.lineEdit_11.setEnabled(True)
            else:
                self.comboBox_27.setEnabled(False)
                self.lineEdit_11.clear()
                self.lineEdit_11.setEnabled(False)

        if type_ == '按钮功能':
            # 切换frame
            self.comboBox_26.activated.connect(switch_frame)
        elif type_ == '写入参数':
            # 切换类型
            parameter_1 = self.comboBox_26.currentText()
            # 获取frame类型
            parameter_2 = self.comboBox_27.currentText().replace('：', '')
            # 获取frame
            parameter_3 = self.lineEdit_11.text()
            if parameter_1 == '切换回主文档' or parameter_1 == '切换到上一级文档':
                parameter_2 = None
                parameter_3 = None
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))

    def save_form_function(self, type_):
        """保存网页表格的功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            parameter_2 = None
            # 获取元素类型和元素
            image = self.comboBox_28.currentText().replace('：', '') + '-' + self.lineEdit_12.text()
            # 获取Excel工作簿路径和工作表名称
            parameter_1 = self.comboBox_29.currentText() + "-" + self.lineEdit_13.text()
            # 判断其他参数
            if self.radioButton_13.isChecked() and not self.radioButton_12.isChecked():
                parameter_2 = '自动跳过'
            elif not self.radioButton_13.isChecked() and self.radioButton_12.isChecked():
                parameter_2 = self.spinBox_9.value()
            # 检查参数是否有异常
            if not self.lineEdit_12.text():
                QMessageBox.critical(self, "错误", "元素未填写！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 image_=image,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            self.comboBox_29.clear()
            self.comboBox_29.addItems(extract_excel_from_global_parameter())

    def drag_element_function(self, type_):
        """拖动网页元素的功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            parameter_2 = None
            # 获取元素类型和元素
            image = self.comboBox_30.currentText().replace('：', '') + '-' + self.lineEdit_14.text()
            # 获取拖动距离
            parameter_1 = str(self.spinBox_10.value()) + "-" + str(self.spinBox_11.value())
            # 判断其他参数
            if self.radioButton_15.isChecked() and not self.radioButton_14.isChecked():
                parameter_2 = '自动跳过'
            elif not self.radioButton_15.isChecked() and self.radioButton_14.isChecked():
                parameter_2 = self.spinBox_12.value()
            # 判断参数是否有异常
            if not self.lineEdit_14.text():
                QMessageBox.critical(self, "错误", "元素未填写！")
                raise ValueError
            if parameter_1 == '0-0':
                QMessageBox.critical(self, "错误", "拖动距离未设置！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 image_=image,
                                                 remarks_=func_info_dic.get('备注'))

    def full_screen_capture_function(self, type_):
        """全屏截图的窗口功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            folder_path = self.comboBox_31.currentText()
            image_name = self.lineEdit_16.text()
            # 检查参数是否有异常
            if image_name == '':
                QMessageBox.critical(self, "错误", "图像名称未填！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=folder_path + '/' + image_name,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            self.comboBox_31.clear()
            self.comboBox_31.addItems(extract_global_parameter('资源文件夹路径'))

    def switch_window_function(self, type_):
        """切换浏览器窗口的功能"""
        if type_ == '按钮功能':
            pass
        elif type_ == '写入参数':
            # 获取窗口类型
            parameter_1 = self.comboBox_32.currentText().replace('：', '')
            # 获取窗口
            parameter_2 = self.lineEdit_15.text()
            # 检查参数是否有异常
            if parameter_2 == '':
                QMessageBox.critical(self, "错误", "窗口未设置！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def wechat_function(self, type_):
        """微信发送消息的功能"""

        def Lock_control():
            """锁定控件"""
            if self.comboBox_33.currentText() == '自定义联系人':
                self.lineEdit_17.setEnabled(True)
            else:
                self.lineEdit_17.setEnabled(False)
                self.lineEdit_17.clear()

            if self.comboBox_34.currentText() == '自定义消息内容':
                self.textEdit_2.setEnabled(True)
            else:
                self.textEdit_2.setEnabled(False)
                self.textEdit_2.clear()

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.comboBox_33.currentText() \
                if self.comboBox_33.currentText() == '文件传输助手' else self.lineEdit_17.text()
            parameter_2_ = self.comboBox_34.currentText() \
                if self.comboBox_34.currentText() != '自定义消息内容' else self.textEdit_2.toPlainText()
            if parameter_1_ == '' or parameter_2_ == '':
                QMessageBox.critical(self, "错误", "联系人或消息内容不能为空！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                '联系人': parameter_1_,
                '消息内容': parameter_2_,
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            contact = parameter_dic_.get('联系人', '')
            message = parameter_dic_.get('消息内容', '')
            # 设置联系人
            if contact == '文件传输助手':
                self.comboBox_33.setCurrentText(contact)
                self.lineEdit_17.setEnabled(False)
            else:
                self.comboBox_33.setCurrentText('自定义联系人')
                self.lineEdit_17.setEnabled(True)
                self.lineEdit_17.setText(contact)
            # 设置消息内容
            if message == '自定义消息内容':
                self.comboBox_34.setCurrentText(message)
                self.textEdit_2.setEnabled(False)
            else:
                self.comboBox_34.setCurrentText('自定义消息内容')
                self.textEdit_2.setEnabled(True)
                self.textEdit_2.setText(message)

        def test():
            """测试"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_dic_)
                # 测试用例
                test_class = SendWeChat(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            Lock_control()
            self.comboBox_33.activated.connect(Lock_control)
            self.comboBox_34.activated.connect(Lock_control)
            self.pushButton_15.clicked.connect(test)
            self.pushButton_30.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))

        elif type_ == '写入参数':
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_dic,
                                                 remarks_=func_info_dic.get('备注'))

        elif type_ == '还原参数':
            put_parameters(self.parameter_1)

    def verification_code_function(self, type_):
        """数字验证码功能"""

        def test():
            """测试功能"""
            try:
                image_, parameter_1_ = get_parameters(True)
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_,
                                         parameter_1_=parameter_1_
                                         )
                # 测试用例
                verification_code = VerificationCode(self.out_mes, dic_)
                verification_code.is_test = True
                verification_code.start_execute()
            except Exception as e:
                self.out_mes.out_mes(f'识别失败，错误信息：{type(e)}', True)

        def set_region():
            """设置区域"""
            screen_capture = ScreenCapture()
            screen_capture.screenshot_area()
            self.label_85.setText(str(screen_capture.region))

        def open_setting_window():
            """打开图像点击设置窗口"""
            setting_win = Setting(self)  # 设置窗体
            setting_win.tabWidget.setCurrentIndex(1)  # 切换到第2页
            setting_win.setModal(True)
            setting_win.exec_()

        def get_parameters(is_test: bool = False):
            """从tab页获取参数"""
            image_ = self.lineEdit_18.text()  # 输入框元素的定位
            parameter_1_ = self.label_85.text()  # 截图区域
            parameter_2_ = self.comboBox_25.currentText().replace('：', '')  # 元素类型
            parameter_3_ = self.comboBox_62.currentText()  # 验证码类型
            # 检查参数是否有异常
            if image_ == '' and not is_test:
                QMessageBox.critical(self, "错误", "元素未设置！")
                raise ValueError
            if parameter_1_ == '(0,0,0,0)':
                QMessageBox.critical(self, "错误", "验证码识别区域未设置！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                '区域': parameter_1_,
                '元素类型': parameter_2_,
                '验证码类型': parameter_3_,
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到tab页"""
            # 设置输入框元素的定位
            self.lineEdit_18.setText(image_)
            # 设置截图区域
            self.label_85.setText(parameter_dic_['区域'])
            # 设置元素类型
            index = self.comboBox_25.findText(parameter_dic_['元素类型'] + '：')
            if index >= 0:
                self.comboBox_25.setCurrentIndex(index)
            # 设置验证码类型
            index = self.comboBox_62.findText(parameter_dic_['验证码类型'])
            if index >= 0:
                self.comboBox_62.setCurrentIndex(index)

        if type_ == '按钮功能':
            self.pushButton_16.clicked.connect(set_region)
            self.pushButton_53.clicked.connect(open_setting_window)
            # 测试按钮
            self.pushButton_17.clicked.connect(test)
        elif type_ == '写入参数':
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_dic,
                                                 remarks_=func_info_dic.get('备注'))

        elif type_ == '还原参数':
            put_parameters(self.image_path, self.parameter_1)

    def play_voice_function(self, type_):
        """播放语音的功能"""

        def get_parameters():
            parameter_1_ = None
            parameter_2_ = None
            parameter_3_ = None
            if self.radioButton_8.isChecked():
                parameter_1_ = '音频信号'
                parameter_2_ = (f'{self.spinBox_21.value()},{self.spinBox_23.value()},'
                                f'{self.spinBox_22.value()},{self.spinBox_24.value()}')  # 信号频率
            elif self.radioButton_9.isChecked():
                parameter_1_ = '系统提示音'
                parameter_2_ = self.comboBox_7.currentText()
            elif self.radioButton_20.isChecked():
                parameter_1_ = '播放语音'
                parameter_2_ = self.textEdit_4.toPlainText()
                parameter_3_ = self.horizontalSlider.value()
            # 检查参数是否有异常
            if self.radioButton_20.isChecked() and parameter_2_ == '':
                QMessageBox.critical(self, "错误", "内容未输入！")
                raise ValueError
            return parameter_1_, parameter_2_, parameter_3_

        def test():
            """测试功能"""
            try:
                parameter_1_, parameter_2_, parameter_3_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_,
                                         parameter_3_=parameter_3_)
                play_voice = PlayVoice(self.out_mes, dic_)
                play_voice.is_test = True
                play_voice.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'参数异常', True)

        if type_ == '按钮功能':
            # 将不同的单选按钮添加到同一个按钮组
            buttonGroup_4 = QButtonGroup(self)
            buttonGroup_4.addButton(self.radioButton_8)
            buttonGroup_4.addButton(self.radioButton_9)
            buttonGroup_4.addButton(self.radioButton_20)
            # 测试按钮
            self.pushButton_24.clicked.connect(test)
            self.pushButton_32.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))

        elif type_ == '写入参数':
            parameter_1, parameter_2, parameter_3 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))

    def wait_window_function(self, type_):
        """倒计时等待窗口的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = f'{self.lineEdit_2.text()}' or '示例'
            parameter_2_ = f'{self.lineEdit_6.text()}' or '示例'
            parameter_3_ = f'{self.spinBox_25.value()}'
            # # 检查参数是否有异常
            return parameter_1_, parameter_2_, parameter_3_

        def test():
            # """测试功能"""
            parameter_1_, parameter_2_, parameter_3_ = get_parameters()
            dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                     parameter_1_=parameter_1_,
                                     parameter_2_=parameter_2_,
                                     parameter_3_=parameter_3_)

            # 测试用例
            test_class = WaitWindow(self.out_mes, dic_)
            test_class.is_test = True
            test_class.start_execute()

        if type_ == '按钮功能':
            self.pushButton_25.clicked.connect(test)
        elif type_ == '写入参数':
            parameter_1, parameter_2, parameter_3 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))

    def dialog_window_function(self, type_):
        """xxx的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.lineEdit_8.text() or '提示框'  # 提示框标题
            parameter_2_ = self.lineEdit_20.text() or '示例'  # 提示框内容
            parameter_3_ = self.comboBox_36.currentText()  # icon类型
            return parameter_1_, parameter_2_, parameter_3_

        def test():
            """测试功能"""
            try:
                parameter_1_, parameter_2_, parameter_3_ = get_parameters()
                dic_ = self.get_test_dic(parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_,
                                         parameter_3_=parameter_3_,
                                         repeat_number_=int(self.spinBox.value())
                                         )

                # 测试用例
                test_class = DialogWindow(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误！', True)

        if type_ == '按钮功能':
            self.pushButton_26.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_1, parameter_2, parameter_3 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))

    def branch_jump_function(self, type_):
        """跳转分支的功能
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.comboBox_37.currentText()  # 分支表名
            parameter_2_ = self.comboBox_38.currentText()  # 分支序号
            return parameter_1_, parameter_2_

        if type_ == '按钮功能':
            # self.comboBox_37.activated.connect(set_branch_count)
            self.comboBox_37.activated.connect(lambda: self.find_controls('分支', '跳转分支'))

        elif type_ == '写入参数':
            parameter_1, parameter_2 = get_parameters()
            # 检查参数是否有异常
            if parameter_1 == '' or parameter_2 == '':
                QMessageBox.critical(self, "错误", "分支为空，请先添加！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=f'{parameter_1}-{parameter_2}',
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            self.comboBox_37.addItems(extract_global_parameter('分支表名'))
            self.comboBox_37.setCurrentIndex(0)
            # 获取分支表名中的指令数量
            self.find_controls('分支', '跳转分支')

    def termination_process_function(self, type_):
        """终止流程的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.comboBox_39.currentText()
            return parameter_1_

        if type_ == '按钮功能':
            pass

        elif type_ == '写入参数':
            exception_handling_ = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=exception_handling_,
                                                 remarks_=func_info_dic.get('备注'))

    def window_control_function(self, type_):
        """窗口控制的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.lineEdit_21.text()
            parameter_2_ = f'{self.comboBox_40.currentText()}-{self.checkBox_5.isChecked()}'
            # 检查参数是否有异常
            if parameter_1_ == '':
                QMessageBox.critical(self, "错误", "窗口标题未填！")
                raise ValueError
            return parameter_1_, parameter_2_

        def test():
            """测试功能"""
            try:
                parameter_1_, parameter_2_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_)

                # 测试用例
                test_class = WindowControl(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                # self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_27.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_1, parameter_2 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def key_wait_function(self, type_):
        """按键等待的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_2_ = None
            parameter_3_ = None
            parameter_1_ = self.keySequenceEdit_2.keySequence().toString()
            if self.radioButton_22.isChecked():
                parameter_2_ = '等待按键'
                parameter_3_ = '自动跳过'
            elif self.radioButton_21.isChecked():
                parameter_2_ = '等待跳转分支'
                parameter_3_ = f'{self.comboBox_41.currentText()}-{self.comboBox_42.currentText()}'

            # 检查参数是否有异常
            if parameter_1_ == '':
                QMessageBox.critical(self, "错误", "按键未设置！")
                raise ValueError
            if parameter_1_.count('+') >= 1:
                QMessageBox.critical(self, "错误", "该功能暂不支持复合按键！")
                raise ValueError
            if parameter_1_.lower() in ['esc', 'f10', 'f11']:
                QMessageBox.critical(self, "错误", "该功能暂不支持设置为esc、f10、f11键！")
                raise ValueError
            if self.radioButton_21.isChecked() and (
                    self.comboBox_41.currentText() == '' or self.comboBox_42.currentText() == ''
            ):
                QMessageBox.critical(self, "错误", "分支异常，请先添加！")
                raise ValueError
            return parameter_1_, parameter_2_, parameter_3_

        def set_branch_name():
            """当选择跳转分支功能时，加载分支表名"""
            disable_control(True)
            self.comboBox_41.addItems(extract_global_parameter('分支表名'))

        def disable_control(judge_: bool):
            """禁用控件"""
            self.comboBox_41.clear()
            self.comboBox_42.clear()
            self.label_133.setEnabled(judge_)
            self.label_132.setEnabled(judge_)
            self.comboBox_41.setEnabled(judge_)
            self.comboBox_42.setEnabled(judge_)

        if type_ == '按钮功能':
            self.radioButton_21.toggled.connect(set_branch_name)
            self.radioButton_22.toggled.connect(lambda: disable_control(False))
            self.comboBox_41.activated.connect(lambda: self.find_controls('分支', '按键等待'))

        elif type_ == '写入参数':
            parameter_1, parameter_2, exception_handling_ = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=exception_handling_,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def gain_time_function(self, type_):
        """获取时间的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.comboBox_43.currentText()
            parameter_2_ = self.comboBox_44.currentText()
            # 检查参数是否有异常
            if parameter_2_ == '':
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            return parameter_1_, parameter_2_

        def test():
            """测试功能"""
            try:
                parameter_1_, parameter_2_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_)

                # 测试用例
                test_class = GetTimeValue(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_33.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))
            self.pushButton_34.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_1, parameter_2 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_44.clear()
            self.comboBox_44.addItems(get_variable_info('list'))

    def gain_excel_function(self, type_):
        """从excel单元格中获取变量的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = f'{self.comboBox_45.currentText()}-{self.comboBox_46.currentText()}'  # Excel路径-工作表
            parameter_1_ = self.lineEdit_23.text()  # 单元格
            parameter_2_ = self.comboBox_47.currentText()  # 变量
            parameter_3_ = str(self.checkBox_9.isChecked())  # 是否行号递增
            # 检查参数是否有异常
            if self.comboBox_45.currentText() == '' or self.comboBox_46.currentText() == '':
                QMessageBox.critical(self, "错误", "Excel路径未设置！")
                raise ValueError
            if self.lineEdit_23.text() == '':
                QMessageBox.critical(self, "错误", "单元格未设置！")
                raise ValueError
            if self.comboBox_47.currentText() == '':
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError

            return image_, parameter_1_, parameter_2_, parameter_3_

        def test():
            """测试功能"""
            try:
                image_, parameter_1_, parameter_2_, parameter_3_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_,
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_,
                                         parameter_3_=parameter_3_)

                # 测试用例
                test_class = GetExcelCellValue(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        def line_number_increasing():
            # 行号递增功能被选中后弹出提示框
            if self.checkBox_9.isChecked():
                QMessageBox.information(self, '提示',
                                        '启用该功能后，请在主页面中设置循环次数大于1，执行全部指令后，'
                                        '循环执行时，单元格行号会自动递增。',
                                        QMessageBox.Ok
                                        )

        if type_ == '按钮功能':
            # 禁用中文输入
            self.lineEdit_23.setValidator(QRegExpValidator(QRegExp("[a-zA-Z0-9]{16}"), self))
            self.checkBox_9.clicked.connect(line_number_increasing)
            self.comboBox_45.activated.connect(
                lambda: self.find_controls('excel', '获取Excel')
            )
            self.pushButton_35.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))
            self.pushButton_36.clicked.connect(test)
            # 打开工作簿
            self.pushButton_29.clicked.connect(lambda: os.startfile(self.comboBox_45.currentText()))

        elif type_ == '写入参数':
            image, parameter_1, parameter_2, parameter_3 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_45.clear()
            self.comboBox_45.addItems(extract_excel_from_global_parameter())  # 加载全局参数中的excel文件路径
            self.find_controls('excel', '获取Excel')

            self.comboBox_47.clear()
            self.comboBox_47.addItems(get_variable_info('list'))

    def get_dialog_function(self, type_):
        """从对话框中获取变量的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.lineEdit_24.text()  # 输入框标题
            parameter_2_ = self.comboBox_48.currentText()  # 变量名称
            parameter_3_ = self.lineEdit_25.text()  # 提示信息
            # 检查参数是否有异常
            if parameter_1_ == '':
                parameter_1_ = '示例'
            if parameter_3_ == '':
                parameter_3_ = '示例'
            if parameter_2_ == '':
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError

            return parameter_1_, parameter_2_, parameter_3_

        if type_ == '按钮功能':
            self.pushButton_37.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))

        elif type_ == '写入参数':
            parameter_1, parameter_2, parameter_3, = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_48.clear()
            self.comboBox_48.addItems(get_variable_info('list'))

    def contrast_variables_function(self, type_):
        """变量比较的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = (
                f'{self.comboBox_49.currentText()}-{self.comboBox_51.currentText()}'
            )  # 变量1-变量2
            parameter_2_ = (
                f'{self.comboBox_50.currentText()}-{self.comboBox_55.currentText()}'
            )  # 比较符-变量类型
            exception_handling_ = (
                f'{self.comboBox_52.currentText()}'
                f'-{self.comboBox_53.currentText()}'
            )  # 分支表名-分支序号
            # 检查参数是否有异常
            if (self.comboBox_49.currentText() == ''
                    or self.comboBox_50.currentText() == ''
                    or self.comboBox_51.currentText() == ''):
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            if (self.comboBox_52.currentText() == ''
                    or self.comboBox_53.currentText() == ''):
                QMessageBox.critical(self, "错误", "分支未设置！")
                raise ValueError
            return parameter_1_, parameter_2_, exception_handling_

        def sync_combo_boxes(sender):
            if sender == self.comboBox_54:
                self.comboBox_55.setCurrentIndex(self.comboBox_54.currentIndex())
            else:
                self.comboBox_54.setCurrentIndex(self.comboBox_55.currentIndex())

        if type_ == '按钮功能':
            self.comboBox_52.activated.connect(  # 当分支表名改变时，加载分支中的命令序号
                lambda: self.find_controls('分支', '变量判断'))
            self.pushButton_38.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))
            self.comboBox_54.currentIndexChanged.connect(lambda: sync_combo_boxes(self.comboBox_54))
            self.comboBox_55.currentIndexChanged.connect(lambda: sync_combo_boxes(self.comboBox_55))

        elif type_ == '写入参数':
            parameter_1, parameter_2, exception_handling = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=exception_handling,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_49.clear()
            self.comboBox_49.addItems(get_variable_info('list'))
            self.comboBox_51.clear()
            self.comboBox_51.addItems(get_variable_info('list'))
            self.comboBox_52.clear()
            self.comboBox_52.addItems(extract_global_parameter('分支表名'))
            self.comboBox_52.setCurrentIndex(0)
            # 获取分支表名中的指令数量
            self.find_controls('分支', '变量判断')

    def run_python_function(self, type_):
        """运行python代码的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.lineEdit_26.text()  # 返回名称
            parameter_2_ = self.comboBox_56.currentText()  # 变量名称
            parameter_3_ = self.textEdit_5.toPlainText()  # 代码
            # 检查参数是否有异常
            if parameter_3_ == '':
                QMessageBox.critical(self, "错误", "代码未编写！")
                raise ValueError
            return parameter_1_, parameter_2_, parameter_3_

        def test():
            """测试功能"""
            highlight_python_code()
            try:
                parameter_1_, parameter_2_, parameter_3_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_,
                                         parameter_3_=parameter_3_)
                # 测试用例
                test_class = RunPython(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        def highlight_python_code():
            """运行python代码"""

            def highlight_text(text):
                lexer = PythonLexer()
                formatter = HtmlFormatter(style='monokai')
                html = highlight(text, lexer, formatter)
                css = formatter.get_style_defs('.highlight')
                self.textEdit_5.setHtml("<style>" + css + "</style>" + html)

            code = self.textEdit_5.toPlainText()
            highlight_text(code)

        if type_ == '按钮功能':
            # 自动代码高亮
            self.pushButton_40.clicked.connect(test)
            self.pushButton_39.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))
            self.pushButton_41.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))

        elif type_ == '写入参数':
            parameter_1, parameter_2, parameter_3 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 parameter_3_=parameter_3,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_56.clear()
            self.comboBox_56.addItems(get_variable_info('list'))

    def run_external_file_function(self, type_):
        """运行外部文件的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = self.lineEdit_27.text()  # 文件路径
            # 检查参数是否有异常
            if image_ is None or image_ == '':
                QMessageBox.critical(self, "错误", "文件路径未设置！")
                raise ValueError
            return image_

        def test():
            """测试功能"""
            try:
                image_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_)

                # 测试用例
                test_class = RunExternalFile(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)

        def get_file_and_folder():
            """获取文件名和文件夹路径"""
            # 打开选择文件对话框
            file_path, _ = QFileDialog.getOpenFileName(
                parent=self,
                caption="选择文件",
                directory=os.path.join(os.path.expanduser("~"), 'Desktop'),
            )
            if file_path != '':  # 获取文件名称
                # 设置文件路径
                self.lineEdit_27.setText(os.path.normpath(file_path))

        if type_ == '按钮功能':
            self.pushButton_43.clicked.connect(get_file_and_folder)  # 打开文件选择窗口
            self.pushButton_42.clicked.connect(test)  # 测试按钮

        elif type_ == '写入参数':
            image = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            pass

    def input_cell_function(self, type_):
        """输入到excel单元格的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = f'{self.comboBox_57.currentText()}-{self.comboBox_58.currentText()}'  # Excel路径-工作表
            parameter_1_ = f'{self.lineEdit_28.text()}-{self.checkBox_10.isChecked()}'  # 单元格-是否行号递增
            parameter_2_ = f'{self.textEdit_6.toPlainText()}'  # 内容
            # 检查参数是否有异常
            if image_ == '':
                QMessageBox.critical(self, "错误", "Excel路径未设置！")
                raise ValueError
            if parameter_1_ == '':
                QMessageBox.critical(self, "错误", "单元格未设置！")
                raise ValueError
            if parameter_2_ == '':
                QMessageBox.critical(self, "错误", "输出内容未设置！")
                raise ValueError

            return image_, parameter_1_, parameter_2_

        if type_ == '按钮功能':
            self.pushButton_44.clicked.connect(lambda: os.startfile(self.comboBox_57.currentText()))
            self.pushButton_45.clicked.connect(lambda: self.merge_additional_functions('打开变量选择'))
            # 禁用中文输入
            self.lineEdit_28.setValidator(QRegExpValidator(QRegExp("[a-zA-Z0-9]{16}"), self))
            self.comboBox_57.activated.connect(
                lambda: self.find_controls('excel', '写入单元格')
            )

        elif type_ == '写入参数':
            image, parameter_1, parameter_2 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_57.clear()
            self.comboBox_57.addItems(extract_excel_from_global_parameter())
            self.find_controls('excel', '写入单元格')

    def ocr_recognition_function(self, type_):
        """ocr的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.label_153.text()  # 识别区域
            parameter_2_ = self.comboBox_59.currentText()  # 写入变量
            # 检查参数是否有异常
            if parameter_1_ == '(0,0,0,0)' or parameter_2_ == '':
                QMessageBox.warning(self, '警告', '参数不能为空！')
                raise Exception
            return parameter_1_, parameter_2_

        def open_setting_window():
            """打开图像点击设置窗口"""
            setting_win = Setting(self)  # 设置窗体
            setting_win.tabWidget.setCurrentIndex(1)  # 切换到第2页
            setting_win.setModal(True)
            setting_win.exec_()

        def set_the_screenshot_area():
            """设置截图区域"""
            screen_capture = ScreenCapture()
            screen_capture.screenshot_area()
            self.label_153.setText(str(screen_capture.region))
            # 显示区域边框
            self.transparent_window.setGeometry(*screen_capture.region)
            self.transparent_window.show()

        def test():
            """测试功能"""
            try:
                parameter_1_, parameter_2_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_)

                # 测试用例
                client_info = get_ocr_info()
                if client_info['appId'] != '':
                    test_class = TextRecognition(self.out_mes, dic_)
                    test_class.is_test = True
                    test_class.start_execute()
                else:
                    QMessageBox.warning(self, '提示', 'OCR未设置！')
                    open_setting_window()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_48.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))
            self.pushButton_46.clicked.connect(set_the_screenshot_area)
            self.pushButton_49.clicked.connect(open_setting_window)  # 打开百度ocr设置
            self.pushButton_47.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_1, parameter_2 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_59.clear()
            self.comboBox_59.addItems(get_variable_info('list'))

    def get_mouse_position_function(self, type_):
        """获取鼠标位置的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_dic = {
                '变量': self.comboBox_60.currentText()
            }
            return parameter_dic

        if type_ == '按钮功能':
            self.pushButton_51.clicked.connect(lambda: self.merge_additional_functions('打开变量池'))

        elif type_ == '写入参数':
            parameter_1 = str(get_parameters())
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_60.clear()
            self.comboBox_60.addItems(get_variable_info('list'))
