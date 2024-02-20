import datetime
import io
import os
import random
import re
import sqlite3

import ddddocr
import openpyxl
import pyautogui
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices, QImage, QPixmap, QIntValidator
from PyQt5.QtWidgets import QWidget, \
    QMessageBox, QButtonGroup, QApplication
from dateutil.parser import parse
from openpyxl.utils.exceptions import InvalidFileException

from 功能类 import SendWeChat, ImageClick, OutputMessage, CoordinateClick, PlayVoice, WaitWindow, DialogWindow
from 截图模块 import ScreenCapture
from 数据库操作 import extract_global_parameter, extract_excel_from_global_parameter, get_branch_count, \
    sqlitedb, close_database, set_window_size, save_window_size
from 窗体.navigation import Ui_navigation
from 网页操作 import WebOption
from 设置窗口 import Setting


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
        # 添加保存按钮事件
        self.modify_id = None
        self.modify_row = None
        self.pushButton_2.clicked.connect(lambda: self.save_data())
        # 获取鼠标位置参数
        self.mouse_position_function = None
        # 调整异常处理选项时，控制窗口控件的状态
        self.comboBox_9.activated.connect(lambda: self.exception_handling_judgment_type('报错处理'))
        self.comboBox_10.activated.connect(lambda: self.exception_handling_judgment_type('分支名称'))
        self.combo_image_preview = {
            '图像点击': (self.comboBox_8, self.comboBox),
            '图像等待': (self.comboBox_17, self.comboBox_18),
            '信息录入': (self.comboBox_14, self.comboBox_15),
        }
        self.combo_excel_preview = {
            '信息录入': (self.comboBox_12, self.comboBox_13),
            '网页录入': (self.comboBox_20, self.comboBox_23),
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
        }
        # 加载功能窗口的按钮功能
        for func_name in self.function_mapping:
            self.function_mapping[func_name][0]('按钮功能')
            self.function_mapping[func_name][0]('加载信息')
        self.tabWidget_2.setCurrentIndex(0)  # 设置到功能页面到预览页
        # 设置窗口的flag
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
        self.main_window.get_data(self.modify_row)
        # 窗口大小
        save_window_size((self.width(), self.height()), self.windowTitle())

    def switch_navigation_page(self, name):
        """弹出窗口自动选择对应功能页
        :param name: 功能页名称"""
        # print('选择功能页：', name)
        try:
            tab_index = self.tab_title_list.index(name)
            self.tabWidget.setCurrentIndex(tab_index)
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
        self.tabWidget_2.setCurrentIndex(3)
        return {
            'ID': None,
            '图像路径': image_,
            '参数1（键鼠指令）': parameter_1_,
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
            # 检查参数是否有异常
            if (os.path.isdir(image_)) or (not os.path.exists(image_)):
                QMessageBox.critical(self, "错误", "图像文件不存在，请检查图像文件是否存在！")
                raise FileNotFoundError
            return image_, parameter_1_, parameter_2_

        def test():
            """测试功能"""
            try:
                image_, parameter_1_, parameter_2_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_,
                                         parameter_1_=parameter_1_,
                                         parameter_2_=parameter_2_,
                                         parameter_3_=str(self.checkBox.isChecked()),  # 灰度识别
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
            # 加载下拉列表数据
            self.comboBox_8.currentTextChanged.connect(
                lambda: self.find_images('图像点击')
            )
            # 元素预览
            self.comboBox.currentTextChanged.connect(
                lambda: self.show_image_to_label(self.comboBox_8, self.comboBox)
            )
            # 测试按钮
            self.pushButton_6.clicked.connect(test)
            # 打开设置窗口
            self.pushButton_11.clicked.connect(open_setting_window)

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
                                                 parameter_3_=str(self.checkBox.isChecked()),  # 灰度识别
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 加载图像文件夹路径
            self.comboBox_8.clear()
            self.comboBox_8.addItems(extract_global_parameter('资源文件夹路径'))

    def scroll_wheel_function(self, type_):
        """滚轮滑动的窗口功能"""
        if type_ == '按钮功能':
            # 将不同的单选按钮添加到同一个按钮组
            buttonGroup_3 = QButtonGroup(self)
            buttonGroup_3.addButton(self.radioButton_gun)
            buttonGroup_3.addButton(self.radioButton_sv)
            self.lineEdit_3.setValidator(QIntValidator())  # 设置只能输入数字
        elif type_ == '写入参数':
            parameter_1 = None
            parameter_2 = None
            if self.radioButton_gun.isChecked():
                parameter_1 = '滚轮滑动'
                parameter_2 = f'{self.comboBox_5.currentText()},{self.lineEdit_3.text()}'  # 鼠标滚轮滑动的方向
            elif self.radioButton_sv.isChecked():
                parameter_1 = '随机滚轮滑动'
                parameter_2 = f'{self.spinBox_16.value()},{self.spinBox_17.value()}'
            # 检查参数是否有异常
            if not self.lineEdit_3.text().isdigit() and self.radioButton_gun.isChecked():
                QMessageBox.critical(self, "错误", "滚动的距离未输入！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 remarks_=func_info_dic.get('备注'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2)

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
            self.comboBox_3.currentTextChanged.connect(spinBox_2_enable)
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
            self.comboBox_17.currentTextChanged.connect(
                lambda: self.find_images('图像等待')
            )
            # 元素预览
            self.comboBox_18.currentTextChanged.connect(
                lambda: self.show_image_to_label(self.comboBox_17, self.comboBox_18)
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

    def move_mouse_function(self, type_):
        """鼠标移动识别窗口的功能"""
        if type_ == '按钮功能':
            # 将不同的单选按钮添加到同一个按钮组
            buttonGroup_2 = QButtonGroup(self)
            buttonGroup_2.addButton(self.radioButton_19)
            buttonGroup_2.addButton(self.radioButton_ran)
            # 限制输入框只能输入数字
            self.lineEdit.setValidator(QIntValidator())
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
            elif self.radioButton_2.isChecked():
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
            self.comboBox_12.currentTextChanged.connect(
                lambda: self.find_excel_sheet_name('信息录入')
            )
            # 加载下拉列表数据
            self.comboBox_14.currentTextChanged.connect(
                lambda: self.find_images('信息录入')
            )
            # 图像预览
            self.comboBox_15.currentTextChanged.connect(
                lambda: self.show_image_to_label(self.comboBox_14, self.comboBox_15)
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

    def open_web_page_function(self, type_):
        """打开网址的窗口功能"""

        def web_functional_testing(judge):
            """网页连接测试"""
            if judge == '测试':
                url = self.lineEdit_19.text()
                web_option = WebOption(self.out_mes)
                web_option.web_open_test(url)

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
                    web_option = WebOption(self.out_mes)
                    web_option.install_browser_driver()
                    QMessageBox.information(self, '提示', '浏览器驱动安装完成！', QMessageBox.Yes)

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
        """网页控制的窗口功能"""

        def Lock_control():
            """锁定控件"""
            if self.comboBox_22.currentText() == '输入内容':
                self.textEdit_3.setEnabled(True)
            else:
                self.textEdit_3.clear()
                self.textEdit_3.setEnabled(False)

        if type_ == '按钮功能':
            Lock_control()
            self.comboBox_22.currentTextChanged.connect(Lock_control)
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
            self.comboBox_20.currentTextChanged.connect(
                lambda: self.find_excel_sheet_name('网页录入')
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
            self.comboBox_26.currentTextChanged.connect(switch_frame)
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
            if image_name == '':
                QMessageBox.critical(self, "错误", "图像名称未填！")
                return
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

        def test():
            """测试"""
            # 设置到功能页面的测试页
            self.tabWidget_2.setCurrentIndex(3)
            # 获取联系人
            if self.comboBox_33.currentText() == '自定义联系人':
                parameter_1_ = self.lineEdit_17.text()
            else:
                parameter_1_ = self.comboBox_33.currentText()
            # 获取消息内容
            if self.comboBox_34.currentText() == '自定义消息内容':
                parameter_2_ = self.textEdit_2.toPlainText()
            else:
                parameter_2_ = self.comboBox_34.currentText()
            # 测试
            ins_dic = {
                '参数1（键鼠指令）': parameter_1_,
                '参数2': parameter_2_,
            }
            wechat_option = SendWeChat(self.out_mes, ins_dic)
            wechat_option.is_test = True
            wechat_option.send_message_to_wechat(parameter_1_, parameter_2_, int(self.spinBox.value()))

        if type_ == '按钮功能':
            Lock_control()
            self.comboBox_33.currentTextChanged.connect(Lock_control)
            self.comboBox_34.currentTextChanged.connect(Lock_control)
            self.pushButton_15.clicked.connect(test)
        elif type_ == '写入参数':
            parameter_1 = self.comboBox_33.currentText() \
                if self.comboBox_33.currentText() == '文件传输助手' else self.lineEdit_17.text()
            parameter_2 = self.comboBox_34.currentText() \
                if self.comboBox_34.currentText() != '自定义消息内容' else self.textEdit_2.toPlainText()
            if parameter_1 == '' or parameter_2 == '':
                QMessageBox.critical(self, "错误", "联系人或消息内容不能为空！")
                return
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

    def verification_code_function(self, type_):
        """数字验证码功能"""

        def test():
            """测试"""
            # 设置到功能页面的测试页
            self.tabWidget_2.setCurrentIndex(3)
            region = eval(self.label_85.text())
            if region == (0, 0, 0, 0):
                self.textBrowser.append('请先设置区域！')
            else:
                im = pyautogui.screenshot(region=(region[0], region[1], region[2], region[3]))
                im_bytes = io.BytesIO()
                im.save(im_bytes, format='PNG')
                im_b = im_bytes.getvalue()
                ocr = ddddocr.DdddOcr()
                res = ocr.classification(im_b)
                self.textBrowser.append('识别出的验证码为：' + res)
                # 释放资源
                del im
                del im_bytes

        def set_region():
            """设置区域"""
            screen_capture = ScreenCapture()
            screen_capture.screenshot_area()
            self.label_85.setText(str(screen_capture.region))

        if type_ == '按钮功能':
            self.pushButton_16.clicked.connect(set_region)
            # 测试按钮
            self.pushButton_17.clicked.connect(test)
        elif type_ == '写入参数':
            image = self.lineEdit_18.text()
            parameter_1 = self.label_85.text()
            parameter_2 = self.comboBox_25.currentText().replace('：', '')
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_1,
                                                 parameter_2_=parameter_2,
                                                 remarks_=func_info_dic.get('备注'))

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
            parameter_1_ = f'{self.lineEdit_2.text()}'
            parameter_2_ = f'{self.lineEdit_6.text()}'
            parameter_3_ = f'{self.spinBox_25.value()}'
            # 检查参数是否有异常
            if parameter_1_ == '' or parameter_2_ == '':
                QMessageBox.critical(self, "错误", "信息未填写！")
                raise ValueError
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
            parameter_1_ = self.lineEdit_8.text()  # 对话框标题
            parameter_2_ = self.lineEdit_20.text()  # 对话框内容
            parameter_3_ = self.comboBox_36.currentText()  # 对话框图标
            # 检查参数是否有异常
            if parameter_1_ == '' or parameter_2_ == '':
                QMessageBox.critical(self, "错误", "信息未填写！")
                raise ValueError
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

        def set_branch_count():
            """当分支表名改变时，加载分支中的命令序号"""
            count_record_ = get_branch_count(self.comboBox_37.currentText())
            self.comboBox_38.clear()
            # 加载分支中的命令序号
            branch_order_ = [str(i) for i in range(1, count_record_ + 1)]
            if len(branch_order_) != 0:
                self.comboBox_38.addItems(branch_order_)

        if type_ == '按钮功能':
            self.comboBox_37.currentTextChanged.connect(set_branch_count)

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
            count_record = get_branch_count(self.comboBox_37.currentText())
            # 加载分支中的命令序号
            branch_order = [str(i) for i in range(1, count_record + 1)]
            if len(branch_order) == 0:
                self.comboBox_37.setCurrentIndex(0)
            else:
                self.comboBox_38.addItems(branch_order)

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

    def find_images(self, ins_name: str) -> None:
        """选择图像文件夹并返回文件夹名称
        :param ins_name: 指令名称"""
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

    def find_excel_sheet_name(self, ins_name: str) -> None:
        """获取excel表格中的所有sheet名称
        :param ins_name:指令名称"""
        comboBox_before, comboBox_after = self.combo_excel_preview.get(ins_name)
        excel_path = comboBox_before.currentText()
        try:
            # 用openpyxl获取excel表格中的所有sheet名称
            excel_sheet_name = openpyxl.load_workbook(excel_path).sheetnames
        except FileNotFoundError:
            excel_sheet_name = []
        except InvalidFileException:
            excel_sheet_name = []
        # 清空combox_13中的所有元素
        comboBox_after.clear()
        # 将excel_sheet_name中的所有元素添加到combox_13中
        comboBox_after.addItems(excel_sheet_name)

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

        # 获取当前活动页面的标题
        current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        disable_control = self.function_mapping.get(current_title)[1]
        control_status(disable_control)  # 控制控件的状态
        # 刷新图像预览
        if current_title in self.combo_image_preview.keys():
            self.find_images(current_title)
        if current_title in self.combo_excel_preview.keys():
            self.find_excel_sheet_name(current_title)
        self.tabWidget_2.setCurrentIndex(0)

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
                    # 获取分支表名中的指令数量
                    count_record = get_branch_count(self.comboBox_10.currentText())
                    # 加载分支中的命令序号
                    branch_order = [str(i) for i in range(1, count_record + 1)]
                    if len(branch_order) == 0:
                        self.comboBox_10.setCurrentIndex(0)
                    else:
                        self.comboBox_11.addItems(branch_order)
            elif type_ == '分支名称':  # 分支表名下拉列表变化触发
                count_record = get_branch_count(self.comboBox_10.currentText())
                self.comboBox_11.clear()
                # 加载分支中的命令序号
                branch_order = [str(i) for i in range(1, count_record + 1)]
                if len(branch_order) == 0:
                    QMessageBox.warning(self, '警告', '该分支下没有指令，请先添加！', QMessageBox.Yes)
                else:
                    self.comboBox_11.addItems(branch_order)
        except sqlite3.OperationalError:
            pass

    def quick_screenshot(self, combox_folder, judge):
        """截图功能
        :param combox_folder: 图像文件夹下拉列表
        :param judge: 功能选择（快捷截图、打开文件夹）"""
        if judge == '快捷截图':
            if combox_folder.currentText() == '':
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
                self.find_images(self.tabWidget.tabText(self.tabWidget.currentIndex()))

        elif judge == '打开文件夹':
            if combox_folder.currentText() != '':
                os.startfile(os.path.normpath(combox_folder.currentText()))

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
                image_, instruction_, parameter_1_, parameter_2_, parameter_3_, parameter_4_, repeat_number_,
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
        func_selected = self.function_mapping.get(tab_title)[0]  # 获取当前页的功能
        # 根据功能获取参数
        if func_selected:
            try:
                func_selected('写入参数')
                self.close()
            except Exception as e:
                print(e)

    def show_image_to_label(self, comboBox_folder, comboBox_image, judge='显示'):
        """将图像显示到label中,图像预览的功能
        :param judge: 显示、删除、查看
        :param comboBox_folder: 图像文件夹下拉列表
        :param comboBox_image: 图像名称下拉列表"""
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

    def on_button_clicked(self, judge: str) -> None:
        """按钮点击事件,用于图像预览的按钮事件
        :param judge: 执行的操作(删除、查看)"""
        # 获取当前页的标题
        current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        if current_title in self.combo_image_preview:
            combo_box1, combo_box2 = self.combo_image_preview.get(current_title)
            self.show_image_to_label(combo_box1, combo_box2, judge)
            if judge == '删除':
                for value in self.combo_image_preview.values():
                    self.find_images(current_title)
                    if value[1].currentText() == '':
                        self.label_43.setText('暂无')
