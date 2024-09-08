import datetime
import glob
import os
import re
import sqlite3

import pyautogui
from PyQt5 import QtCore
from PyQt5.QtCore import QUrl, QRegExp, Qt
from PyQt5.QtGui import (
    QDesktopServices,
    QImage,
    QPixmap,
    QIntValidator,
    QRegExpValidator, QKeySequence,
)
from PyQt5.QtWidgets import (
    QMessageBox,
    QTreeWidgetItemIterator,
    QFileDialog,
    QWidget,
    QApplication, QDialog, QColorDialog,
)
from dateutil.parser import parse
from openpyxl.utils.exceptions import InvalidFileException
from pandas import ExcelFile
from pygments import highlight
from pygments.formatters import HtmlFormatter
from pygments.lexers import PythonLexer

from ini控制 import (
    set_window_size,
    save_window_size,
    extract_resource_folder_path,
    get_branch_info,
    get_ocr_info,
    get_all_png_images_from_resource_folders,
    matched_complete_path_from_resource_folders
)
from 功能类 import (
    InformationEntry,
    InputCellExcel,
    MouseDrag,
    OutputMessage,
    TransparentWindow,
    ImageClick,
    CoordinateClick,
    PlayVoice,
    WaitWindow,
    DialogWindow,
    WindowControl,
    GetTimeValue,
    GetExcelCellValue,
    RunPython,
    RunExternalFile,
    TextRecognition,
    VerificationCode,
    SendWeChat,
    MoveMouse,
    FullScreenCapture,
    MultipleImagesClick,
    RunCmd,
    GetClipboard,
    ColorJudgment,
)
from 变量池窗口 import VariablePool_Win
from 图像点击位置 import ClickPosition
from 截图模块 import ScreenCapture
from 数据库操作 import (
    extract_excel_from_global_parameter,
    get_branch_count,
    sqlitedb,
    close_database,
    get_variable_info
)
from 窗体.图像选择 import Ui_ImageSelect
from 窗体.导航窗口 import Ui_navigation
from 网页操作 import WebOption
from 设置窗口 import Setting
from 选择窗体 import Variable_selection_win, ShortcutTable


class ImageSelection(QDialog, Ui_ImageSelect):
    """选择图片窗口"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )  # 隐藏帮助按钮
        self.load_images_name_to_listView()  # 加载图片名称到listView
        self.listView.clicked.connect(self.preview_image)  # 预览图片
        self.pushButton.clicked.connect(self.get_image_name)  # 获取图片名称
        self.pushButton_2.clicked.connect(self.close)  # 关闭窗口

    def load_images_name_to_listView(self):
        """加载图片名称到listView"""
        images_name_list = get_all_png_images_from_resource_folders()
        # 创建模型并绑定到listView
        model = QtCore.QStringListModel()
        model.setStringList(images_name_list)
        self.listView.setModel(model)

    def preview_image(self):
        """预览图片"""
        # 获取图片名称
        image_name = self.listView.currentIndex().data()
        # 获取图片路径
        image_path = matched_complete_path_from_resource_folders(image_name)
        # 加载图片
        if image_path != '':
            # 将图像转换为QImage对象
            image_ = QImage(image_path)
            image = image_.scaled(
                self.label.width(),
                self.label.height(),
                Qt.KeepAspectRatio,
            )
            self.label.setPixmap(QPixmap.fromImage(image))

    def get_image_name(self):
        """获取图片路径"""
        image_name = self.listView.currentIndex().data()
        print('选择的图片名称:', image_name)
        try:
            # 获取父窗口中的listView和其模型
            parent_listView = self.parent().listView
            model = parent_listView.model()
            # 如果模型不存在，就创建一个新的 QStringListModel
            if model is None:
                model = QtCore.QStringListModel()
                parent_listView.setModel(model)
            # 获取当前的字符串列表
            current_list = model.stringList()
            # 检查是否已经存在相同的图片名称
            if image_name in current_list:
                QMessageBox.warning(self, '警告', '该图像不能重复添加！')
            else:
                # 将新的图片名称添加到列表中
                current_list.append(image_name)
                # 更新模型的数据
                model.setStringList(current_list)
                self.close()
        except Exception as e:
            print(e)


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
        self.variable_sel_win = Variable_selection_win(self, "变量选择")  # 变量选择窗口
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
        self.comboBox_9.activated.connect(
            lambda: self.exception_handling_judgment_type("报错处理")
        )
        self.comboBox_10.activated.connect(
            lambda: self.exception_handling_judgment_type("分支名称")
        )
        self.combo_image_preview = {  # 需要图像预览功能
            "图像点击": (self.comboBox_8, self.comboBox),
            "图像等待": (self.comboBox_17, self.comboBox_18),
            "信息录入": (self.comboBox_14, self.comboBox_15),
        }
        self.combo_excel_preview = {  # 需要加载excel表格的功能
            "信息录入": (self.comboBox_12, self.comboBox_13),
            "网页录入": (self.comboBox_20, self.comboBox_23),
            "获取Excel": (self.comboBox_45, self.comboBox_46),
            "写入单元格": (self.comboBox_57, self.comboBox_58),
        }
        self.variable_input_control = {  # 需要插入变量的控件
            "文本输入": self.textEdit,
            "元素控制": self.textEdit_3,
            "发送消息": self.textEdit_2,
            "提示音": self.textEdit_4,
            "运行Python": self.textEdit_5,
            "写入单元格": self.textEdit_6,
            "运行cmd": self.textEdit_7,
        }
        self.branch_jump_control = {  # 需要分支跳转的功能
            "功能区参数": (self.comboBox_10, self.comboBox_11),
            "跳转分支": (self.comboBox_37, self.comboBox_38),
            "变量判断": (self.comboBox_52, self.comboBox_53),
            "按键等待": (self.comboBox_41, self.comboBox_42),
            "颜色判断": (self.comboBox_74, self.comboBox_75),
        }
        self.pushButton_9.clicked.connect(lambda: self.on_button_clicked("查看"))
        self.pushButton_10.clicked.connect(lambda: self.on_button_clicked("删除"))
        # 快捷选择导航页
        self.tab_title_list = [
            self.tabWidget.tabText(x) for x in range(self.tabWidget.count())
        ]
        self.treeWidget.itemClicked.connect(
            lambda: self.switch_navigation_page(self.treeWidget.currentItem().text(0))
        )
        self.tabWidget.currentChanged.connect(self.tab_widget_change)
        # 映射标签标题和对应的函数
        self.function_mapping = {
            "图像点击": (lambda x: self.image_click_function(x), True),
            "多图点击": (lambda x: self.multiple_images_click_function(x), True),
            "坐标点击": (lambda x: self.coordinate_click_function(x), False),
            "移动鼠标": (lambda x: self.move_mouse_function(x), False),
            "时间等待": (lambda x: self.time_waiting_function(x), False),
            "图像等待": (lambda x: self.image_waiting_function(x), True),
            "滚轮滑动": (lambda x: self.scroll_wheel_function(x), False),
            "文本输入": (lambda x: self.text_input_function(x), False),
            "按下键盘": (lambda x: self.press_keyboard_function(x), False),
            "中键激活": (lambda x: self.middle_activation_function(x), False),
            "鼠标点击": (lambda x: self.mouse_click_function(x), False),
            "鼠标拖拽": (lambda x: self.mouse_drag_function(x), False),
            "信息录入": (lambda x: self.information_entry_function(x), True),
            "打开网址": (lambda x: self.open_web_page_function(x), False),
            "元素控制": (lambda x: self.ele_control_function(x), True),
            "网页录入": (lambda x: self.web_entry_function(x), True),
            "切换frame": (lambda x: self.toggle_frame_function(x), False),
            "保存表格": (lambda x: self.save_form_function(x), True),
            "拖动元素": (lambda x: self.drag_element_function(x), True),
            "屏幕截图": (lambda x: self.full_screen_capture_function(x), False),
            "切换窗口": (lambda x: self.switch_window_function(x), False),
            "发送消息": (lambda x: self.wechat_function(x), False),
            "数字验证码": (lambda x: self.verification_code_function(x), True),
            "提示音": (lambda x: self.play_voice_function(x), False),
            "倒计时窗口": (lambda x: self.wait_window_function(x), False),
            "提示窗口": (lambda x: self.dialog_window_function(x), False),
            "跳转分支": (lambda x: self.branch_jump_function(x), False),
            "终止流程": (lambda x: self.termination_process_function(x), False),
            "窗口控制": (lambda x: self.window_control_function(x), True),
            "按键等待": (lambda x: self.key_wait_function(x), False),
            "获取时间": (lambda x: self.gain_time_function(x), False),
            "获取Excel": (lambda x: self.gain_excel_function(x), False),
            "获取对话框": (lambda x: self.get_dialog_function(x), False),
            "获取剪切板": (lambda x: self.get_clipboard_function(x), False),
            "变量判断": (lambda x: self.contrast_variables_function(x), False),
            "运行Python": (lambda x: self.run_python_function(x), False),
            "运行cmd": (lambda x: self.run_cmd_function(x), False),
            "运行外部文件": (lambda x: self.run_external_file_function(x), True),
            "写入单元格": (lambda x: self.input_cell_function(x), True),
            "OCR识别": (lambda x: self.ocr_recognition_function(x), False),
            "获取鼠标位置": (lambda x: self.get_mouse_position_function(x), False),
            "窗口焦点等待": (lambda x: self.window_focus_wait_function(x), True),
            "颜色判断": (lambda x: self.color_judgment_function(x), False),
        }
        # 加载功能窗口的按钮功能
        for func_name in self.function_mapping:
            self.function_mapping[func_name][0]("按钮功能")
        # 加载第一个功能窗口的控件信息
        self.function_mapping[self.tabWidget.tabText(0)][0]("加载信息")
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
        self.disable_exception_handling_control(False)  # 禁用异常处理控件

    def closeEvent(self, a0) -> None:
        """关闭窗口时,触发的动作"""
        if self.transparent_window.isVisible():  # 关闭框选窗口
            self.transparent_window.close()
        self.main_window.get_data(self.modify_row)
        # 窗口大小
        save_window_size(self.width(), self.height(), self.windowTitle())

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
            if exception_handling_text in {
                "自动跳过",
                "提示异常并暂停",
                "提示异常并停止",
            }:
                self.comboBox_9.setCurrentText(exception_handling_text)
            elif "-" in exception_handling_text:
                # 处理跳转分支的情况
                select_branch_table_name, branch_index = exception_handling_text.split(
                    "-"
                )
                self.comboBox_9.setCurrentText("跳转分支")
                # 解除异常处理方式的禁用，加载分支表名
                self.comboBox_10.addItems(get_branch_info(True))
                self.find_controls("分支", "功能区参数")
                # self.comboBox_10.setEnabled(True)
                # self.comboBox_11.setEnabled(True)
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
                    reverse_exception_handling_judgment(
                        restore_parameters[3]
                    )  # 恢复异常处理参数
                    func_selected("还原参数")
                    self.show_image_to_label(name)  # 显示图像
                except Exception as e:
                    QMessageBox.warning(self, "警告", "该指令参数错误", QMessageBox.Yes)
                    print("还原参数错误", e)

        except ValueError:  # 如果没有找到对应的功能页，则跳过
            pass

    def get_test_dic(
            self,
            repeat_number_,
            image_=None,
            parameter_1_=None,
            parameter_2_=None,
            parameter_3_=None,
            parameter_4_=None,
    ):
        """返回测试字典,用于测试按钮的功能"""
        self.tabWidget_2.setCurrentIndex(2)
        return {
            "ID": None,
            "图像路径": image_,
            "参数1（键鼠指令）": str(parameter_1_),
            "参数2": parameter_2_,
            "参数3": parameter_3_,
            "参数4": parameter_4_,
            "重复次数": repeat_number_,
        }

    def get_func_info(self) -> dict:
        """返回功能区的参数"""

        def exception_handling_judgment():
            """判断异常处理方式
            :return: 异常处理方式"""
            exception_handling_text = None
            selected_text = self.comboBox_9.currentText()
            if selected_text in {"自动跳过", "提示异常并暂停", "提示异常并停止"}:
                exception_handling_text = selected_text
            elif selected_text == "跳转分支":
                select_branch_table_name = self.comboBox_10.currentText()
                if self.comboBox_11.currentText() == "":
                    QMessageBox.critical(
                        self, "错误", "分支表下无指令，请检查分支表名是否正确！"
                    )
                    raise ValueError
                exception_handling_text = (
                    f"{select_branch_table_name}-{int(self.comboBox_11.currentText())}"
                )
            return exception_handling_text

        # 当前页的index
        tab_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
        return {
            "重复次数": self.spinBox.value(),
            "异常处理": exception_handling_judgment(),
            "备注": self.lineEdit_5.text(),
            "指令类型": tab_title,
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
                images_name = [f for f in os.listdir(folder_path) if f.endswith(".png")]
                # Sort files by modification time
                images_name.sort(
                    key=lambda x: os.path.getmtime(os.path.join(folder_path, x)),
                    reverse=True,
                )
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
                excel_sheet_name = ExcelFile(excel_path).sheet_names
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
            comboBox_branch_name, comboBox_branch_order = self.branch_jump_control.get(
                ins_name
            )
            count_record_ = get_branch_count(comboBox_branch_name.currentText())
            comboBox_branch_order.clear()
            # 加载分支中的命令序号
            branch_order_ = [str(i) for i in range(1, count_record_ + 1)]
            if len(branch_order_) != 0:
                comboBox_branch_order.addItems(branch_order_)

        if type_ == "图像":
            find_images()
        elif type_ == "excel":
            find_excel_sheet_name()
        elif type_ == "分支":
            find_branch_count()

    def mouseMoveEvent(self, event):
        self.merge_additional_functions("get_mouse_position")

    def tab_widget_change(self):
        """切换导航页功能"""

        def control_status(disable_control_):
            """控制控件的状态，功能区参数控件的状态"""
            self.label_33.setVisible(disable_control_)
            self.label_34.setVisible(disable_control_)
            self.label_35.setVisible(disable_control_)
            self.comboBox_9.setCurrentIndex(0)
            self.comboBox_9.setVisible(disable_control_)
            self.disable_exception_handling_control(False)

        try:
            # 获取当前活动页面的标题
            current_title = self.tabWidget.tabText(self.tabWidget.currentIndex())
            disable_control = self.function_mapping.get(current_title)[1]
            control_status(disable_control)  # 控制控件的状态
            if self.transparent_window.isVisible():  # 关闭框选窗口
                self.transparent_window.close()
                # 加载功能窗口的按钮功能
            self.function_mapping[current_title][0]("加载信息")
            self.tabWidget_2.setCurrentIndex(0)
        except TypeError:
            pass

    def merge_additional_functions(self, function_name, pars_1=None):
        """将一次性和冗余的功能合并
        :param pars_1:参数1
        :param function_name: 功能名称
        """
        def get_rgb_value():
            """获取颜色的rgb值"""
            # 获取鼠标位置的rgb值
            rgb = pyautogui.pixel(x, y)
            self.spinBox_26.setValue(rgb[0])
            self.spinBox_29.setValue(rgb[1])
            self.spinBox_30.setValue(rgb[2])
            # 设置标签的背景色
            self.label_191.setStyleSheet(
                f"background-color:rgb({rgb[0]},{rgb[1]},{rgb[2]})"
            )

        if function_name == "get_mouse_position":
            # 获取鼠标位置
            x, y = pyautogui.position()
            if self.mouse_position_function == "坐标点击":
                self.label_9.setText(str(x))
                self.label_10.setText(str(y))
            elif self.mouse_position_function == "开始拖拽":
                self.label_59.setText(str(x))
                self.label_61.setText(str(y))
            elif self.mouse_position_function == "结束拖拽":
                self.label_65.setText(str(x))
                self.label_66.setText(str(y))
            elif self.mouse_position_function == "指定坐标":
                self.lineEdit_29.setText(str(x))
                self.lineEdit_30.setText(str(y))
            elif self.mouse_position_function == "颜色判断":
                self.label_197.setText(str(x))
                self.label_195.setText(str(y))
                # 获取鼠标位置的rgb值
                get_rgb_value()
            elif self.mouse_position_function == "获取颜色":
                # 获取鼠标位置的rgb值
                get_rgb_value()
        elif function_name == "change_get_mouse_position_function":
            # 改变获取鼠标位置功能
            self.mouse_position_function = pars_1
        elif function_name == "打开变量池":
            variable_pool = VariablePool_Win(self)
            variable_pool.exec_()
        elif function_name == "打开变量选择":
            self.variable_sel_win.show()

    def disable_exception_handling_control(self, judge: bool = False):
        """禁用控件"""
        self.comboBox_10.clear()
        self.label_34.setVisible(judge)
        self.comboBox_10.setVisible(judge)
        self.comboBox_11.clear()
        self.label_35.setVisible(judge)
        self.comboBox_11.setVisible(judge)

    def exception_handling_judgment_type(self, type_):
        """判断异常护理选项并调整控件
        :param type_: 判断类型（报错处理、分支名称）"""

        try:
            if type_ == "报错处理":  # 报错处理下拉列表变化触发
                if self.comboBox_9.currentText() == "自动跳过":
                    self.disable_exception_handling_control(False)
                elif self.comboBox_9.currentText() == "提示异常并暂停":
                    self.disable_exception_handling_control(False)
                elif self.comboBox_9.currentText() == "提示异常并停止":
                    self.disable_exception_handling_control(False)
                elif self.comboBox_9.currentText() == "跳转分支":
                    self.disable_exception_handling_control(True)
                    self.comboBox_10.addItems(get_branch_info(True))
                    self.comboBox_10.setCurrentIndex(0)
                    self.find_controls("分支", "功能区参数")
            elif type_ == "分支名称":  # 分支表名下拉列表变化触发
                self.find_controls("分支", "功能区参数")
        except sqlite3.OperationalError:
            pass

    def quick_screenshot(self, control_name, judge):
        """截图功能
        :param control_name: 需要的控件
        :param judge: 功能选择（快捷截图、打开文件夹）"""
        if judge == "快捷截图":
            if control_name.currentText() == "":
                QMessageBox.warning(self, "警告", "未选择图像文件夹！", QMessageBox.Yes)
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
                self.find_controls("图像", current_title)
                self.show_image_to_label(current_title)

        elif judge == "打开文件夹":
            if control_name.currentText() != "":
                os.startfile(os.path.normpath(control_name.currentText()))
            else:
                QMessageBox.warning(
                    self, "警告", "未设置资源文件夹，请前往主页设置！", QMessageBox.Yes
                )

        elif judge == "设置区域":
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

    def writes_commands_to_the_database(
            self,
            instruction_,
            repeat_number_,
            exception_handling_,
            image_=None,
            parameter_1_=None,
            parameter_2_=None,
            parameter_3_=None,
            parameter_4_=None,
            remarks_=None,
    ):
        """向数据库写入命令"""
        try:
            cursor, con = sqlitedb()
            branch_name = self.main_window.comboBox.currentText()

            query_params = (
                image_,
                instruction_,
                str(parameter_1_),
                parameter_2_,
                parameter_3_,
                parameter_4_,
                repeat_number_,
                exception_handling_,
                remarks_,
                branch_name,
            )
            if self.pushButton_2.text() == "添加指令":
                cursor.execute(
                    "INSERT INTO 命令"
                    "(图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)",
                    query_params,
                )

            elif self.pushButton_2.text() == "修改指令":
                cursor.execute(
                    "UPDATE 命令 "
                    "SET 图像名称=?,指令类型=?,参数1=?,参数2=?,参数3=?,参数4=?,重复次数=?,异常处理=?,备注=?,隶属分支=? "
                    "WHERE ID=?",
                    query_params + (self.modify_id,),
                )

            elif self.pushButton_2.text() == "向前插入":
                # 将当前ID和之后的ID递增1
                max_id_ = 1000000
                cursor.execute(
                    "UPDATE 命令 SET ID=ID+? WHERE ID>=?", (max_id_, self.modify_id)
                )
                cursor.execute(
                    "UPDATE 命令 SET ID=ID-? WHERE ID>=?",
                    (max_id_ - 1, max_id_ + int(self.modify_id)),
                )
                # 插入新的命令
                cursor.execute(
                    "INSERT INTO 命令"
                    "(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (self.modify_id,) + query_params,
                )

            elif self.pushButton_2.text() == "向后插入":
                self.modify_row = self.modify_row + 1
                try:
                    cursor.execute(
                        "INSERT INTO 命令"
                        "(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        (self.modify_id + 1,) + query_params,
                    )
                except sqlite3.IntegrityError:
                    # 如果下一个id已经存在，则将后面的id全部加1
                    max_id_ = 1000000
                    cursor.execute(
                        "UPDATE 命令 SET ID=ID+? WHERE ID>?", (max_id_, self.modify_id)
                    )
                    cursor.execute(
                        "UPDATE 命令 SET ID=ID-? WHERE ID>?",
                        (max_id_ - 1, max_id_ + int(self.modify_id)),
                    )
                    # 插入新的命令
                    cursor.execute(
                        "INSERT INTO 命令"
                        "(ID,图像名称,指令类型,参数1,参数2,参数3,参数4,重复次数,异常处理,备注,隶属分支) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        (self.modify_id + 1,) + query_params,
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
                    func_selected("写入参数")
                    self.close()
                except Exception as e:
                    print(e)
        except TypeError:
            pass

    def show_image_to_label(self, ins_name: str, judge="显示"):
        """将图像显示到label中,图像预览的功能
        :param ins_name: 指令名称
        :param judge: 显示、删除、查看"""
        try:
            comboBox_folder, comboBox_image = self.combo_image_preview.get(ins_name)
            image_path = os.path.normpath(
                os.path.join(
                    comboBox_folder.currentText(), comboBox_image.currentText()
                )
            )
            if (os.path.exists(image_path)) and (
                    os.path.isfile(image_path)
            ):  # 判断图像是否存在
                if judge == "显示":
                    # 将图像转换为QImage对象
                    image_ = QImage(image_path)
                    image = image_.scaled(
                        self.label_43.width(),
                        self.label_43.height(),
                        Qt.KeepAspectRatio,
                    )
                    self.label_43.setPixmap(QPixmap.fromImage(image))
                elif judge == "删除":
                    os.remove(image_path)
                elif judge == "查看":
                    os.startfile(image_path)
            else:
                self.label_43.setText("暂无")
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
            if judge == "删除":
                self.find_controls("图像", current_title)
                self.show_image_to_label(current_title)

    def write_value_to_textedit(self, value: str) -> None:
        """将变量池中的值写入到文本框中"""

        def append_textedit(new_text):
            errorFormat_ = '<font color="red">{}</font>'
            # 使textEdit显示不同的文本
            current_title_ = self.tabWidget.tabText(self.tabWidget.currentIndex())
            textEdit = self.variable_input_control.get(current_title_)
            if textEdit.isEnabled():
                textEdit.insertHtml("☾")
                textEdit.insertHtml((errorFormat_.format(new_text)))
                textEdit.insertHtml("☽")

        if value:
            append_textedit(value)

    def image_click_function(self, type_):
        """图像点击识别窗口的功能
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = os.path.normpath(
                os.path.join(self.comboBox_8.currentText(), self.comboBox.currentText())
            )
            parameter_1_ = self.comboBox_2.currentText()
            # 如果复选框被选中，则获取第二个参数
            parameter_2_ = None
            if self.radioButton_2.isChecked():
                parameter_2_ = "自动略过"
            elif self.radioButton_4.isChecked():
                parameter_2_ = self.spinBox_4.value()
            if self.groupBox_57.isChecked():
                parameter_4_ = self.label_155.text()
            else:
                parameter_4_ = "(0,0,0,0)"
            # 从tab页获取参数
            parameter_dic_ = {
                "动作": parameter_1_,
                "异常": parameter_2_,
                "区域": parameter_4_,
                "灰度": self.checkBox.isChecked(),
                "精度": self.horizontalSlider_4.value() / 100,
                "点击位置": self.label_176.text(),
            }
            # 检查参数是否有异常
            if (os.path.isdir(image_)) or (not os.path.exists(image_)):
                QMessageBox.critical(
                    self, "错误", "图像文件不存在，请检查图像文件是否存在！"
                )
                raise FileNotFoundError
            if self.groupBox_57.isChecked() and self.label_155.text() == "(0,0,0,0)":
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
            self.find_controls("图像", "图像点击")
            image_file_index = self.comboBox.findText(image_file)
            if image_file_index != -1:
                self.comboBox.setCurrentIndex(image_file_index)
            else:
                # 如果文件不存在，则添加文件
                self.comboBox.addItem(image_file)
                self.comboBox.setCurrentIndex(self.comboBox.findText(image_file))

            # 将其他参数设置回对应的控件
            self.comboBox_2.setCurrentText(parameter_dic_["动作"])

            if parameter_dic_["异常"] == "自动略过":
                self.radioButton_2.setChecked(True)
            else:
                self.radioButton_4.setChecked(True)
                self.spinBox_4.setValue(parameter_dic_["异常"])

            if parameter_dic_["区域"] == "(0,0,0,0)":
                self.groupBox_57.setChecked(False)
            else:
                self.groupBox_57.setChecked(True)
                self.label_155.setText(parameter_dic_["区域"])
            self.checkBox.setChecked(parameter_dic_["灰度"])
            self.horizontalSlider_4.setValue(int(parameter_dic_["精度"] * 100))
            self.label_176.setText(parameter_dic_["点击位置"])
            self.show_image_to_label("图像点击")

        def test():
            """测试功能"""
            try:
                image_, parameter_1_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    image_=image_,
                    parameter_1_=parameter_1_,
                )
                # 测试用例
                try:
                    image_click = ImageClick(self.out_mes, dic_)
                    image_click.is_test = True
                    image_click.start_execute()
                except Exception as e:
                    print(e)
                    self.out_mes.out_mes(f"未找到目标图像，测试结束", True)

            except FileNotFoundError:
                self.out_mes.out_mes(f"图像文件未设置！", True)

        def open_setting_window():
            """打开图像点击设置窗口"""
            setting_win = Setting(self)  # 设置窗体
            setting_win.tabWidget.setCurrentIndex(0)
            setting_win.setModal(True)
            setting_win.exec_()

        def open_set_click_position_window():
            """打开设置点击位置窗口"""
            image_path = os.path.normpath(
                os.path.join(self.comboBox_8.currentText(), self.comboBox.currentText())
            )
            position = self.label_176.text()
            set_click_position = ClickPosition(self, image_path, position)
            set_click_position.setModal(True)
            set_click_position.exec_()

        if type_ == "按钮功能":
            # 快捷截图功能
            self.pushButton.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_8, "快捷截图")
            )
            # 打开图像文件夹
            self.pushButton_7.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_8, "打开文件夹")
            )
            # 设置区域
            self.pushButton_50.clicked.connect(
                lambda: self.quick_screenshot(self.label_155, "设置区域")
            )
            # 加载下拉列表数据
            self.comboBox_8.activated.connect(
                lambda: self.find_controls("图像", "图像点击")
            )
            # 元素预览
            self.comboBox.activated.connect(
                lambda: self.show_image_to_label("图像点击")
            )
            # 测试按钮
            self.pushButton_6.clicked.connect(test)
            # 打开设置窗口
            self.pushButton_11.clicked.connect(open_setting_window)
            # 打开设置点击位置窗口
            self.pushButton_76.clicked.connect(open_set_click_position_window)
            # 设置识别精度
            self.horizontalSlider_4.valueChanged.connect(
                lambda: self.label_178.setText(f'{str(self.horizontalSlider_4.value())}%')
            )

        elif type_ == "写入参数":
            image, parameter_1 = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                parameter_1_=parameter_1,
                remarks_=func_info_dic.get("备注"),
            )

        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

        elif type_ == "加载信息":
            # 加载图像文件夹路径
            self.comboBox_8.clear()
            self.comboBox_8.addItems(extract_resource_folder_path())
            self.find_controls("图像", "图像点击")
            self.show_image_to_label("图像点击")
            # 设置初始识别精度
            self.label_178.setText(f'{str(self.horizontalSlider_4.value())}%')

    def multiple_images_click_function(self, type_):
        """多图像点击识别窗口的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def open_images_select_window():
            """打开图像选择窗口"""
            images_select = ImageSelection(self)
            images_select.setModal(True)
            images_select.exec_()

        def quick_screenshot_and_add_image():
            """快捷截图并添加图像"""

            def get_the_latest_saved_image():
                """获取最新保存的图像"""
                latest_image_path_ = None
                latest_mod_time = 0
                for folder_path in extract_resource_folder_path():
                    for png_file in glob.glob(os.path.join(folder_path, '*.png')):
                        mod_time = os.path.getmtime(png_file)
                        if mod_time > latest_mod_time:
                            latest_image_path_, latest_mod_time = png_file, mod_time
                return latest_image_path_

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
            latest_image_path = get_the_latest_saved_image()
            if latest_image_path:
                # 向listWidget中添加图像名称
                try:
                    image_name = os.path.basename(latest_image_path)
                    # 获取父窗口中的listView和其模型
                    parent_listView = self.listView
                    model = parent_listView.model()
                    # 如果模型不存在，就创建一个新的 QStringListModel
                    if model is None:
                        model = QtCore.QStringListModel()
                        parent_listView.setModel(model)
                    # 获取当前的字符串列表
                    current_list = model.stringList()
                    # 将新的图片名称添加到列表中
                    current_list.append(image_name)
                    # 更新模型的数据
                    model.setStringList(current_list)
                except Exception as e:
                    print(e)

        def delect_selected_image():
            """删除选中的图像"""
            # 获取listView中选中的图像名称
            selected_image = self.listView.currentIndex().data()
            if selected_image:
                # 弹出提示对话框
                reply = QMessageBox.question(
                    self,
                    '删除图像',
                    f'是否删除本地图像：{selected_image}？\n删除后不可恢复。',
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    # 获取图像完整路径
                    image_path = matched_complete_path_from_resource_folders(selected_image)
                    # 刷新listView
                    model = self.listView.model()
                    current_list = model.stringList()
                    current_list.remove(selected_image)
                    model.setStringList(current_list)
                    # 删除图像
                    os.remove(image_path)

        def remove_checked_image():
            """移除选中的图像名称"""
            selected_image = self.listView.currentIndex().data()
            if selected_image:
                model = self.listView.model()
                current_list = model.stringList()
                current_list.remove(selected_image)
                model.setStringList(current_list)

        def move_checked_image(direction):
            """移动选中的图像名称"""
            selected_index = self.listView.currentIndex()
            selected_row = selected_index.row()
            model = self.listView.model()
            current_list = model.stringList()
            if direction == "上移" and selected_row > 0:
                target_row = selected_row - 1
            elif direction == "下移" and selected_row < len(current_list) - 1:
                target_row = selected_row + 1
            else:
                return
            current_list[selected_row], current_list[target_row] = current_list[target_row], current_list[selected_row]
            model.setStringList(current_list)
            self.listView.setCurrentIndex(model.index(target_row))

        def preview_image():
            # 获取当前选中的图像名称
            selected_image = self.listView.currentIndex().data()
            if selected_image:
                # 获取图像完整路径
                image_path = matched_complete_path_from_resource_folders(selected_image)
                # 显示图像
                image_ = QImage(image_path)
                image_ = image_.scaled(
                    self.label_43.width(),
                    self.label_43.height(),
                    Qt.KeepAspectRatio,
                )
                self.label_43.setPixmap(QPixmap.fromImage(image_))
                self.tabWidget_2.setCurrentIndex(1)  # 设置到功能页面到预览页

        def open_setting_window():
            """打开图像点击设置窗口"""
            setting_win = Setting(self)  # 设置窗体
            setting_win.tabWidget.setCurrentIndex(0)
            setting_win.setModal(True)
            setting_win.exec_()

        def get_parameters():
            """从tab页获取参数"""
            images_name_list = self.listView.model().stringList() if self.listView.model() else []
            if not images_name_list:
                QMessageBox.critical(self, "错误", "未添加任何图像！")
                raise FileNotFoundError
            if self.groupBox_73.isChecked():
                if self.label_174.text() == "(0,0,0,0)":
                    QMessageBox.critical(self, "错误", "未设置识别区域！")
                    raise FileNotFoundError
            # 返回参数字典
            image_ = '、'.join(images_name_list)
            parameter_dic_ = {
                '动作': self.comboBox_70.currentText(),
                '异常': self.comboBox_71.currentText()[-4:],
                '区域': self.label_174.text() if self.groupBox_73.isChecked() else "(0,0,0,0)",
                '灰度': self.checkBox_11.isChecked(),
                '精度': self.horizontalSlider_5.value() / 100
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到tab页"""
            # 将图像名称添加到listView中
            image_name_list = image_.split('、')
            model = self.listView.model()
            if model is None:
                model = QtCore.QStringListModel()
                self.listView.setModel(model)
            model.setStringList(image_name_list)
            # 将其他参数设置回对应的控件
            self.comboBox_70.setCurrentText(parameter_dic_['动作'])
            self.comboBox_71.setCurrentIndex(1 if parameter_dic_['异常'] == "自动略过" else 0)
            if parameter_dic_['区域'] == "(0,0,0,0)":
                self.groupBox_73.setChecked(False)
            else:
                self.groupBox_73.setChecked(True)
                self.label_174.setText(parameter_dic_['区域'])
            self.checkBox_11.setChecked(parameter_dic_['灰度'])
            self.horizontalSlider_5.setValue(int(parameter_dic_['精度'] * 100))

        def test():
            """测试功能"""
            try:
                image_, parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         image_=image_,
                                         parameter_1_=parameter_dic_
                                         )

                # 测试用例
                test_class = MultipleImagesClick(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_67.clicked.connect(open_images_select_window)
            # 快捷截图
            self.pushButton_63.clicked.connect(quick_screenshot_and_add_image)
            # 删除选中图像
            self.pushButton_68.clicked.connect(delect_selected_image)
            # 移除选中
            self.pushButton_64.clicked.connect(remove_checked_image)
            # 移动
            self.toolButton_2.clicked.connect(lambda: move_checked_image("上移"))
            self.toolButton.clicked.connect(lambda: move_checked_image("下移"))
            self.listView.clicked.connect(preview_image)
            self.pushButton_66.clicked.connect(open_setting_window)
            # 设置区域
            self.pushButton_65.clicked.connect(
                lambda: self.quick_screenshot(self.label_174, "设置区域")
            )
            self.pushButton_69.clicked.connect(test)
            # 设置识别精度
            self.horizontalSlider_5.valueChanged.connect(
                lambda: self.label_181.setText(f'{str(self.horizontalSlider_5.value())}%')
            )

        elif type_ == '写入参数':
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 image_=image,
                                                 parameter_1_=parameter_dic,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.label_181.setText(f'{str(self.horizontalSlider_5.value())}%')

        elif type_ == '还原参数':
            put_parameters(self.image_path, self.parameter_1)

    def scroll_wheel_function(self, type_):
        """滚轮滑动的窗口功能"""

        def put_parameters(parameter_dic__):
            """将参数还原到窗体控件"""
            parameter_1_ = self.parameter_1["类型"]
            if parameter_1_ == "随机滚轮滑动":
                self.groupBox_29.setChecked(True)
                self.groupBox_22.setChecked(False)
                self.spinBox_16.setValue(parameter_dic__["最小距离"])
                self.spinBox_17.setValue(parameter_dic__["最大距离"])
            elif parameter_1_ == "滚轮滑动":
                self.groupBox_22.setChecked(True)
                self.groupBox_29.setChecked(False)
                self.comboBox_5.setCurrentText(parameter_dic__["方向"])
                self.lineEdit_3.setText(parameter_dic__["距离"])

        if type_ == "按钮功能":
            # 将不同的单选按钮添加到同一个按钮组
            all_groupBoxes_ = [self.groupBox_22, self.groupBox_29]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(
                    lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_)
                )
            self.lineEdit_3.setValidator(QIntValidator())  # 设置只能输入数字

        elif type_ == "写入参数":
            parameter_dic_ = None
            parameter_1 = None
            if self.groupBox_22.isChecked():
                parameter_1 = "滚轮滑动"
            elif self.groupBox_29.isChecked():
                parameter_1 = "随机滚轮滑动"
            # 检查参数是否有异常
            if not self.lineEdit_3.text().isdigit() and self.groupBox_22.isChecked():
                QMessageBox.critical(self, "错误", "滚动的距离未输入！")
                raise ValueError
            # 参数字典
            if parameter_1 == "随机滚轮滑动":
                parameter_dic_ = {
                    "类型": parameter_1,
                    "最小距离": self.spinBox_16.value(),
                    "最大距离": self.spinBox_17.value(),
                }
            elif parameter_1 == "滚轮滑动":
                parameter_dic_ = {
                    "类型": parameter_1,
                    "方向": self.comboBox_5.currentText(),
                    "距离": self.lineEdit_3.text(),
                }
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                remarks_=func_info_dic.get("备注"),
                parameter_1_=parameter_dic_,
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def text_input_function(self, type_):
        """文本输入窗口的功能"""

        def check_text_type():
            # 检查文本输入类型
            text = self.textEdit.toPlainText()
            # 检查text中是否为英文大小写字母和数字
            if (re.search("[a-zA-Z0-9]", text) is None) and (
                    self.checkBox_2.isChecked()
            ):
                self.checkBox_2.setChecked(False)
                QMessageBox.warning(
                    self,
                    "警告",
                    "特殊控件的文本输入仅支持输入英文大小写字母和数字！",
                    QMessageBox.Yes,
                )

        if type_ == "按钮功能":
            # 检查输入的数据是否合法
            self.checkBox_2.clicked.connect(check_text_type)
            self.pushButton_28.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )

        elif type_ == "写入参数":
            # 文本输入的内容
            image = self.textEdit.toPlainText()
            parameter_dic = {"手动输入": str(self.checkBox_2.isChecked())}
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            # 将参数还原到窗体控件
            self.textEdit.setText(self.image_path)
            self.checkBox_2.setChecked(eval(self.parameter_1["手动输入"]))

    def coordinate_click_function(self, type_):
        """坐标点击识别窗口的功能
        :param type_: 功能名称（加载按钮、主要功能）"""

        def spinBox_2_enable() -> None:
            """是否激活自定义点击次数"""
            is_custom = self.comboBox_3.currentText() == "左键（自定义次数）"
            self.spinBox_2.setVisible(is_custom)
            self.label_22.setVisible(is_custom)
            if not is_custom:
                self.spinBox_2.setValue(0)

        def get_parameters():
            """从tab页获取参数"""
            parameter_dic_ = {
                "动作": self.comboBox_3.currentText(),
                "坐标": f"{self.label_9.text()}-{self.label_10.text()}",
                "自定义次数": self.spinBox_2.value(),
            }
            # 检查参数是否有异常
            if self.label_9.text() == "0" and self.label_10.text() == "0":
                QMessageBox.critical(self, "错误", "未设置坐标，请设置坐标！")
                raise ValueError
            return parameter_dic_

        def test():
            """测试功能"""
            parameter_dic_ = get_parameters()
            dic_ = self.get_test_dic(
                repeat_number_=int(self.spinBox.value()), parameter_1_=parameter_dic_
            )
            # 测试用例
            try:
                cor_click = CoordinateClick(self.out_mes, dic_)
                cor_click.is_test = True
                cor_click.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"参数异常", True)

        if type_ == "按钮功能":
            # 坐标点击
            self.pushButton_4.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "坐标点击"
                )
            )
            # 是否激活自定义点击次数
            self.comboBox_3.activated.connect(spinBox_2_enable)
            # 测试按钮
            self.pushButton_23.clicked.connect(test)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )

        elif type_ == "还原参数":
            # 将参数还原到窗体控件
            self.comboBox_3.setCurrentText(self.parameter_1["动作"])
            spinBox_2_enable()
            self.label_9.setText(self.parameter_1["坐标"].split("-")[0])
            self.label_10.setText(self.parameter_1["坐标"].split("-")[1])
            self.spinBox_2.setValue(int(self.parameter_1["自定义次数"]))

    def time_waiting_function(self, type_):
        """等待识别窗口的功能
        :param type_: 功能名称（加载按钮、主要功能）"""

        def time_judgment(target_time):
            """判断时间是否大于当前时间"""
            now_time = datetime.datetime.now()
            return True if now_time < parse(target_time) else False

        def get_now_date_time():
            """将当前的时间和日期设置为dateTimeEdit的日期和时间"""
            if self.groupBox_15.isChecked():
                # 获取当前日期和时间
                now_date_time = datetime.datetime.now().strftime("%H:%M:%S")
                # 将当前的时间和日期加10分钟
                new_date_time = parse(now_date_time) + datetime.timedelta(minutes=10)
                # 将dateTimeEdit的日期和时间设置为当前日期和时间
                self.timeEdit.setDateTime(new_date_time)

        def get_parameters():
            """从tab页获取参数"""
            parameter_dic_ = None
            parameter_1_ = None
            if self.groupBox.isChecked():
                parameter_1_ = "时间等待"
            elif self.groupBox_16.isChecked():
                parameter_1_ = "随机等待"
            elif self.groupBox_15.isChecked():
                parameter_1_ = "定时等待"
            # 参数字典
            if parameter_1_ == "时间等待":
                parameter_dic_ = {
                    "类型": parameter_1_,
                    "时长": self.spinBox_13.value(),
                    "单位": self.comboBox_25.currentText(),
                }
            elif parameter_1_ == "随机等待":
                parameter_dic_ = {
                    "类型": parameter_1_,
                    "最小": f"{self.spinBox_14.value()}-{self.comboBox_64.currentText()}",
                    "最大": f"{self.spinBox_15.value()}-{self.comboBox_65.currentText()}",
                }
            elif parameter_1_ == "定时等待":
                parameter_dic_ = {
                    "类型": parameter_1_,
                    "时间": self.timeEdit.text(),
                    "检测频率": self.comboBox_6.currentText(),
                }
                if not time_judgment(parameter_dic_["时间"]):
                    QMessageBox.critical(
                        self, "错误", "目标时间不能小于当前时间！时间已过。"
                    )
                    raise ValueError
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            if parameter_dic_:
                parameter_type = parameter_dic_.get("类型", None)
                if parameter_type == "时间等待":
                    # 设置 groupBox 的选中状态
                    self.groupBox.setChecked(True)
                    self.groupBox_16.setChecked(False)
                    self.groupBox_15.setChecked(False)
                    # 设置 spinBox_13 和 comboBox_25 的值
                    self.spinBox_13.setValue(parameter_dic_.get("时长", 0))
                    self.comboBox_25.setCurrentText(parameter_dic_.get("单位", "秒"))
                elif parameter_type == "随机等待":
                    # 设置 groupBox_16 的选中状态
                    self.groupBox.setChecked(False)
                    self.groupBox_16.setChecked(True)
                    self.groupBox_15.setChecked(False)
                    # 解析最小和最大值
                    min_value, min_unit = parameter_dic_.get("最小", "0-秒").split("-")
                    max_value, max_unit = parameter_dic_.get("最大", "0-秒").split("-")
                    # 设置 spinBox_14 和 comboBox_64 的值
                    self.spinBox_14.setValue(int(min_value))
                    self.comboBox_64.setCurrentText(min_unit)
                    # 设置 spinBox_15 和 comboBox_65 的值
                    self.spinBox_15.setValue(int(max_value))
                    self.comboBox_65.setCurrentText(max_unit)
                elif parameter_type == "定时等待":
                    # 设置 groupBox_15 的选中状态
                    self.groupBox.setChecked(False)
                    self.groupBox_16.setChecked(False)
                    self.groupBox_15.setChecked(True)
                    # 设置 dateTimeEdit 和 comboBox_6 的值
                    self.timeEdit.setDateTime(
                        QtCore.QDateTime.fromString(
                            parameter_dic_.get("时间", ""), "HH:mm:ss"
                        )
                    )
                    self.comboBox_6.setCurrentText(parameter_dic_.get("检测频率", "秒"))

        if type_ == "按钮功能":
            all_groupBoxes_ = [self.groupBox, self.groupBox_16, self.groupBox_15]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(
                    lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_)
                )
            # 设置当前日期和时间
            self.groupBox_15.clicked.connect(get_now_date_time)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def image_waiting_function(self, type_):
        """图像等待识别窗口的功能"""

        def get_parameters():
            """获取参数"""
            image_ = os.path.normpath(
                os.path.join(
                    self.comboBox_17.currentText(), self.comboBox_18.currentText()
                )
            )
            parameter_1_ = self.comboBox_19.currentText()  # 图像消失类型
            parameter_2_ = self.spinBox_6.value()  # 超时等待时间
            if self.groupBox_61.isChecked():
                parameter_3 = self.label_160.text()  # 区域
            else:
                parameter_3 = "(0,0,0,0)"
            # 从tab页获取参数
            parameter_dic_ = {
                "等待类型": parameter_1_,
                "超时时间": parameter_2_,
                "区域": parameter_3,
                "精度": self.horizontalSlider_6.value() / 100,
            }
            # 检查参数是否有异常
            if (os.path.isdir(image_)) or (not os.path.exists(image_)):
                QMessageBox.critical(
                    self, "错误", "图像文件不存在，请检查图像文件是否存在！"
                )
                raise FileNotFoundError
            if self.groupBox_61.isChecked() and self.label_160.text() == "(0,0,0,0)":
                QMessageBox.critical(self, "错误", "未设置识别区域！")
                raise FileNotFoundError
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原图像路径到comboBox_17和comboBox_18
            image_path, image_file = os.path.split(image_)

            image_index_17 = self.comboBox_17.findText(image_path)
            if image_index_17 != -1:
                self.comboBox_17.setCurrentIndex(image_index_17)
            else:
                self.comboBox_17.addItem(image_path)
                self.comboBox_17.setCurrentIndex(self.comboBox_17.findText(image_path))
            self.find_controls("图像", "图像等待")  # 加载图像
            image_index_18 = self.comboBox_18.findText(image_file)
            if image_index_18 != -1:
                self.comboBox_18.setCurrentIndex(image_index_18)
            else:
                self.comboBox_18.addItem(image_file)
                self.comboBox_18.setCurrentIndex(self.comboBox_18.findText(image_file))
            # 还原等待类型到comboBox_19
            wait_type_index = self.comboBox_19.findText(parameter_dic_["等待类型"])
            self.comboBox_19.setCurrentIndex(wait_type_index)
            # 还原超时时间到spinBox_6
            self.spinBox_6.setValue(parameter_dic_["超时时间"])
            # 还原区域到label_160
            if parameter_dic_["区域"] != "(0,0,0,0)":
                self.groupBox_61.setChecked(True)
                self.label_160.setText(parameter_dic_["区域"])
            else:
                self.groupBox_61.setChecked(False)
                self.label_160.setText("(0,0,0,0)")
            self.horizontalSlider_6.setValue(int(parameter_dic_["精度"] * 100))
            self.show_image_to_label("图像等待")  # 将图像显示到预览中

        if type_ == "按钮功能":
            # 下拉列表数据
            self.comboBox_17.activated.connect(
                lambda: self.find_controls("图像", "图像等待")
            )
            # 元素预览
            self.comboBox_18.activated.connect(
                lambda: self.show_image_to_label("图像等待")
            )
            # 快捷截图功能
            self.pushButton_21.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_17, "快捷截图")
            )
            # 打开图像文件夹
            self.pushButton_22.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_17, "打开文件夹")
            )
            # 设置区域
            self.pushButton_54.clicked.connect(
                lambda: self.quick_screenshot(self.label_160, "设置区域")
            )
            self.horizontalSlider_6.valueChanged.connect(
                lambda: self.label_182.setText(f'{str(self.horizontalSlider_6.value())}%')
            )
        elif type_ == "写入参数":
            # 获取参数
            image, parameter_dic = get_parameters()
            # 检查参数是否有异常
            if (os.path.isdir(image)) or (not os.path.exists(image)):
                QMessageBox.critical(
                    self, "错误", "图像文件不存在，请检查图像文件是否存在！"
                )
                raise FileNotFoundError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                remarks_=func_info_dic.get("备注"),
                image_=image,
                parameter_1_=parameter_dic,
            )
        elif type_ == "加载信息":
            # 加载图像文件夹路径
            self.comboBox_17.clear()
            self.comboBox_18.clear()
            self.comboBox_17.addItems(extract_resource_folder_path())
            self.find_controls("图像", "图像等待")
            self.label_182.setText(f'{str(self.horizontalSlider_6.value())}%')
            self.show_image_to_label("图像等待")

        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def move_mouse_function(self, type_):
        """鼠标移动识别窗口的功能"""

        def get_parameters():
            """获取参数"""

            def check_coordinate(x: int, y: int):
                """检查坐标是否超过屏幕"""
                # 获取屏幕的宽和高
                screen = QApplication.primaryScreen()
                screen_size = screen.size()
                x_, y_ = screen_size.width(), screen_size.height()
                if x > x_ or y > y_:  # 如果坐标超过屏幕范围
                    QMessageBox.critical(
                        self, "错误", f"坐标超过当前屏幕范围{x_, y_}，请重新设置！"
                    )
                    raise ValueError

            parameter_dic_ = None
            # 参数字典
            if self.groupBox_28.isChecked():
                parameter_dic_ = {
                    "类型": "直线移动",
                    "方向": self.comboBox_4.currentText(),
                    "距离": self.lineEdit.text(),
                }
                if not self.lineEdit.text():
                    QMessageBox.critical(self, "错误", "未设置距离！")
                    raise ValueError
            # 参数字典
            if self.groupBox_30.isChecked():
                parameter_dic_ = {
                    "类型": "随机移动",
                    "随机": self.comboBox_16.currentText(),
                }
            elif self.groupBox_59.isChecked():
                parameter_dic_ = {
                    "类型": "指定坐标",
                    "坐标": f"{self.lineEdit_29.text()},{self.lineEdit_30.text()}",
                    "持续": self.doubleSpinBox.value(),
                }
                if not self.lineEdit_29.text() or not self.lineEdit_30.text():
                    QMessageBox.critical(self, "错误", "未设置坐标！")
                    raise ValueError
                # 检查坐标的x和y是超过屏幕
                check_coordinate(
                    int(self.lineEdit_29.text()), int(self.lineEdit_30.text())
                )
            elif self.groupBox_20.isChecked():
                parameter_dic_ = {
                    "类型": "变量坐标",
                    "变量": self.comboBox_61.currentText(),
                }
                if not self.comboBox_61.currentText():
                    QMessageBox.critical(self, "错误", "未设置变量！")
                    raise ValueError
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            if parameter_dic_:
                parameter_type = parameter_dic_.get("类型", None)
                ui_elements = {
                    "直线移动": (
                        self.groupBox_28,
                        self.comboBox_4,
                        self.lineEdit,
                        "方向",
                        "距离",
                    ),
                    "随机移动": (
                        self.groupBox_30,
                        self.comboBox_16,
                        None,
                        "随机",
                        None,
                    ),
                    "指定坐标": (self.groupBox_59, None, None, None, None),
                    "变量坐标": (
                        self.groupBox_20,
                        self.comboBox_61,
                        None,
                        "变量",
                        None,
                    ),
                }
                for key, (
                        group_box,
                        combo_box,
                        line_edit,
                        combo_text,
                        line_text,
                ) in ui_elements.items():
                    group_box.setChecked(parameter_type == key)
                    if combo_box:
                        combo_box.setCurrentText(parameter_dic_.get(combo_text))
                    if line_edit:
                        line_edit.setText(parameter_dic_.get(line_text, "0"))
                if parameter_type == "指定坐标":
                    x, y = parameter_dic_.get("坐标", "0,0").split(",")
                    self.lineEdit_29.setText(x)
                    self.lineEdit_30.setText(y)
                    self.doubleSpinBox.setValue(parameter_dic_.get("持续", 0))

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )

                # 测试用例
                test_class = MoveMouse(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            all_groupBoxes_ = [
                self.groupBox_28,
                self.groupBox_30,
                self.groupBox_59,
                self.groupBox_20,
            ]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(
                    lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_)
                )
            # 自动获取坐标按钮
            self.pushButton_56.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "指定坐标"
                )
            )
            # 限制输入框只能输入数字
            self.lineEdit.setValidator(QIntValidator())
            self.lineEdit_29.setValidator(QIntValidator())
            self.lineEdit_30.setValidator(QIntValidator())
            # 测试按钮
            self.pushButton_52.clicked.connect(test)
            # 打开变量池
            self.pushButton_57.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )

        elif type_ == "加载信息":
            # 加载下拉列表数据
            self.comboBox_61.clear()
            self.comboBox_61.addItems(get_variable_info("list"))

        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def press_keyboard_function(self, type_):
        """按下键盘识别窗口的功能"""

        def method_one(judge):
            """方法一"""
            # 获取“运行Python”标题的索引
            tab_index = self.tab_title_list.index('运行Python')
            self.tabWidget.setCurrentIndex(tab_index)
            code_1 = (
                "import pyautogui\n\n"
                "pyautogui.hotkey('ctrl', 'v')\n"
            )
            code_2 = (
                "import keyboard\n\n"
                "keyboard.press_and_release('ctrl + v')\n"
            )
            self.textEdit_5.setText(code_1 if judge == "pyautogui" else code_2)

        if type_ == "按钮功能":
            # 当按钮按下时，获取按键的名称
            self.pushButton_77.pressed.connect(lambda: method_one("keyboard"))
            self.pushButton_78.pressed.connect(lambda: method_one("pyautogui"))

        elif type_ == "写入参数":
            # 按下键盘的内容
            parameter_dic = {
                "按键": self.keySequenceEdit.keySequence().toString(),
                "按压时长": self.spinBox_27.value(),
            }
            if parameter_dic["按键"] == "":
                QMessageBox.critical(self, "错误", "未设置按键，请设置按键！")
                raise ValueError
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                remarks_=func_info_dic.get("备注"),
                parameter_1_=parameter_dic,
            )
        elif type_ == "还原参数":
            # 将参数还原到窗体控件
            self.keySequenceEdit.setKeySequence(self.parameter_1["按键"])
            self.spinBox_27.setValue(int(self.parameter_1["按压时长"]))

    def middle_activation_function(self, type_):
        """中键激活的窗口功能"""

        def get_parameters():
            """获取参数"""
            parameter_dic_ = None
            if self.radioButton.isChecked():
                parameter_dic_ = {
                    "类型": "模拟点击",
                    "次数": self.spinBox_3.value(),
                }
            elif self.radioButton_zi.isChecked():
                parameter_dic_ = {
                    "类型": "结束等待",
                }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            if parameter_dic_["类型"] == "模拟点击":
                self.radioButton.setChecked(True)
                self.spinBox_3.setValue(parameter_dic_["次数"])
            elif parameter_dic_["类型"] == "结束等待":
                self.radioButton_zi.setChecked(True)

        if type_ == "按钮功能":
            pass
        elif type_ == "写入参数":
            # 中键激活的内容
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def mouse_click_function(self, type_):
        """鼠标点击的窗口的功能"""

        def get_parameters():
            """获取参数"""
            parameter_dic_ = {
                "鼠标": self.comboBox_35.currentText().replace("（自定义次数）", ""),
                "次数": self.spinBox_18.value(),
                "间隔": self.spinBox_20.value(),
                "按压": self.spinBox_19.value(),
            }
            if self.groupBox_63.isChecked():
                parameter_dic_["辅助键"] = self.comboBox_66.currentText()
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            mouse_index = self.comboBox_35.findText(
                parameter_dic_["鼠标"] + "（自定义次数）"
            )
            self.comboBox_35.setCurrentIndex(mouse_index)
            self.spinBox_18.setValue(parameter_dic_["次数"])
            self.spinBox_20.setValue(parameter_dic_["间隔"])
            self.spinBox_19.setValue(parameter_dic_["按压"])
            self.horizontalSlider_2.setValue(parameter_dic_["按压"])
            self.horizontalSlider_3.setValue(parameter_dic_["间隔"])
            if "辅助键" in parameter_dic_:  # 如果有辅助键
                self.groupBox_63.setChecked(True)
                self.comboBox_66.setCurrentText(parameter_dic_["辅助键"])

        if type_ == "按钮功能":
            pass
        elif type_ == "写入参数":
            # 获取鼠标当前位置的参数
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def information_entry_function(self, type_):
        """信息录入的窗口功能"""

        def line_number_increasing():
            # 行号递增功能被选中后弹出提示框
            if self.checkBox_3.isChecked():
                QMessageBox.information(
                    self,
                    "提示",
                    "启用该功能后，请在主页面中设置循环次数大于1，执行全部指令后，"
                    "循环执行时，单元格行号会自动递增。",
                    QMessageBox.Ok,
                )

        def get_parameters():
            """获取参数"""
            # 检查是否有异常
            if not self.lineEdit_4.text():
                QMessageBox.critical(self, "错误", "未设置单元格，请设置单元格！")
                raise ValueError
            if (
                    self.comboBox_12.currentText() == ""
                    or self.comboBox_13.currentText() == ""
            ):
                QMessageBox.critical(
                    self, "错误", "未设置工作簿或工作表，请设置工作簿或工作表！"
                )
                raise ValueError
            if (
                    self.comboBox_14.currentText() == ""
                    or self.comboBox_15.currentText() == ""
            ):
                QMessageBox.critical(self, "错误", "未设置图像，请设置图像！")
                raise ValueError
            # 图像路径
            image_path = os.path.normpath(
                os.path.join(
                    self.comboBox_14.currentText(), self.comboBox_15.currentText()
                )
            )
            # 异常处理
            exception_ = (
                "自动跳过"
                if self.radioButton_3.isChecked() and not self.radioButton_5.isChecked()
                else self.spinBox_5.value()
            )
            # 参数字典
            parameter_dic_ = {
                "工作簿": self.comboBox_12.currentText(),
                "工作表": self.comboBox_13.currentText(),
                "单元格": self.lineEdit_4.text(),
                "递增": str(self.checkBox_3.isChecked()),
                "模拟输入": str(self.checkBox_4.isChecked()),
                "异常": exception_,
            }
            return parameter_dic_, image_path

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原图像路径
            self.comboBox_14.setCurrentText(os.path.split(image_)[0])
            self.find_controls("图像", "信息录入")
            self.comboBox_15.setCurrentText(os.path.split(image_)[1])
            # 还原工作簿
            self.comboBox_12.setCurrentText(parameter_dic_["工作簿"])
            self.find_controls("excel", "信息录入")
            self.comboBox_13.setCurrentText(parameter_dic_["工作表"])
            self.lineEdit_4.setText(parameter_dic_["单元格"])
            self.checkBox_3.setChecked(eval(parameter_dic_["递增"]))
            self.checkBox_4.setChecked(eval(parameter_dic_["模拟输入"]))
            # 还原异常处理
            if parameter_dic_["异常"] == "自动跳过":
                self.radioButton_3.setChecked(True)
                self.radioButton_5.setChecked(False)
            else:
                self.radioButton_3.setChecked(False)
                self.radioButton_5.setChecked(True)
                self.spinBox_5.setValue(int(parameter_dic_["异常"]))

        def test():
            """测试功能"""
            try:
                parameter_dic_, image_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                    image_=image_,
                )
                # 测试用例
                test_class = InformationEntry(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            # 行号自动递增提示
            self.checkBox_3.clicked.connect(line_number_increasing)
            self.lineEdit_4.setValidator(QRegExpValidator(QRegExp("[A-Za-z0-9]+")))
            # 信息录入页面的快捷截图功能
            self.pushButton_5.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_14, "快捷截图")
            )
            self.pushButton_8.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_14, "打开文件夹")
            )
            # 信息录入窗口的excel功能
            self.comboBox_12.activated.connect(
                lambda: self.find_controls("excel", "信息录入")
            )
            # 加载下拉列表数据
            self.comboBox_14.activated.connect(
                lambda: self.find_controls("图像", "信息录入")
            )
            # 图像预览
            self.comboBox_15.activated.connect(
                lambda: self.show_image_to_label("信息录入")
            )
            # 测试按钮
            self.pushButton_58.clicked.connect(test)
        elif type_ == "写入参数":
            parameter_dic, image = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                image_=image,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 加载文件路径
            self.comboBox_12.clear()
            self.comboBox_12.addItems(
                extract_excel_from_global_parameter()
            )  # 加载全局参数中的excel文件路径
            self.comboBox_13.clear()
            self.comboBox_14.clear()
            self.comboBox_14.addItems(extract_resource_folder_path())
            self.find_controls("图像", "信息录入")
            self.show_image_to_label("信息录入")
            self.find_controls("excel", "信息录入")
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def open_web_page_function(self, type_):
        """打开网址的窗口功能"""

        def web_functional_testing(judge):
            """网页连接测试"""
            if judge == "测试":
                url = self.lineEdit_19.text()
                web_option = WebOption(self.out_mes)
                is_succeed, str_info = web_option.web_open_test(
                    url
                )  # 测试网页是否能打开
                if is_succeed:
                    QMessageBox.information(self, "提示", "连接成功！", QMessageBox.Yes)
                else:
                    QMessageBox.critical(self, "错误", str_info)

            elif judge == "安装浏览器":
                url = "https://google.cn/chrome/"
                QDesktopServices.openUrl(QUrl(url))

            elif judge == "安装浏览器驱动":
                # 弹出选择提示框
                x = QMessageBox.information(
                    self,
                    "提示",
                    "确认下载浏览器驱动？",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if x == QMessageBox.Yes:
                    print("下载浏览器驱动")
                    self.tabWidget_2.setCurrentIndex(2)
                    web_option = WebOption(self.out_mes)
                    web_option.install_browser_driver()

        def put_parameters(image_):
            """将参数还原到窗体控件"""
            self.lineEdit_19.setText(image_)

        if type_ == "按钮功能":
            self.pushButton_18.clicked.connect(lambda: web_functional_testing("测试"))
            self.pushButton_19.clicked.connect(
                lambda: web_functional_testing("安装浏览器")
            )
            self.pushButton_20.clicked.connect(
                lambda: web_functional_testing("安装浏览器驱动")
            )
        elif type_ == "写入参数":
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                remarks_=func_info_dic.get("备注"),
                image_=self.lineEdit_19.text(),
            )

        elif type_ == "还原参数":
            put_parameters(self.image_path)

    def ele_control_function(self, type_):
        """网页元素控制的窗口功能"""

        def Lock_control():
            """锁定控件"""
            if self.comboBox_22.currentText() == "输入内容":
                self.textEdit_3.setEnabled(True)
            else:
                self.textEdit_3.clear()
                self.textEdit_3.setEnabled(False)

        def get_parameters():
            """获取参数"""
            image_ = self.lineEdit_7.text()
            # 判断其他参数
            timeout_type = None
            if self.radioButton_6.isChecked() and not self.radioButton_7.isChecked():
                timeout_type = "自动跳过"
            elif not self.radioButton_6.isChecked() and self.radioButton_7.isChecked():
                timeout_type = self.spinBox_7.value()
            # 获取参数字典
            parameter_dic_ = {
                "元素类型": self.comboBox_21.currentText(),
                "文本": self.textEdit_3.toPlainText(),
                "操作": self.comboBox_22.currentText(),
                "超时类型": timeout_type,
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原图像路径到lineEdit_7
            self.lineEdit_7.setText(image_)
            # 还原元素类型到comboBox_21
            element_type_index = self.comboBox_21.findText(parameter_dic_["元素类型"])
            self.comboBox_21.setCurrentIndex(element_type_index)
            # 还原文本到textEdit_3
            self.textEdit_3.setText(parameter_dic_["文本"])
            # 还原操作到comboBox_22
            operation_index = self.comboBox_22.findText(parameter_dic_["操作"])
            self.comboBox_22.setCurrentIndex(operation_index)
            # 还原超时类型到radioButton_6和radioButton_7
            if self.comboBox_22.currentText() == "输入内容":
                self.textEdit_3.setEnabled(True)
            else:
                self.textEdit_3.clear()
                self.textEdit_3.setEnabled(False)
            # 还原超时类型到radioButton_6和radioButton_7
            if parameter_dic_["超时类型"] == "自动跳过":
                self.radioButton_6.setChecked(True)
                self.radioButton_7.setChecked(False)
            else:
                self.radioButton_6.setChecked(False)
                self.radioButton_7.setChecked(True)
                self.spinBox_7.setValue(parameter_dic_["超时类型"])

        if type_ == "按钮功能":
            Lock_control()
            self.comboBox_22.activated.connect(Lock_control)
            self.pushButton_31.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )

        elif type_ == "写入参数":
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                remarks_=func_info_dic.get("备注"),
                image_=image,
                parameter_1_=parameter_dic,
            )
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def web_entry_function(self, type_):
        """网页录入的窗口功能"""

        def get_parameters():
            """获取参数"""
            parameter_4_ = None
            # 判断其他参数
            if self.radioButton_10.isChecked() and not self.radioButton_11.isChecked():
                parameter_4_ = "自动跳过"
            elif (
                    not self.radioButton_10.isChecked() and self.radioButton_11.isChecked()
            ):
                parameter_4_ = self.spinBox_8.value()
            # 获取参数值
            image_path_ = self.comboBox_20.currentText()
            parameter_dic_ = {
                "元素类型": self.comboBox_24.currentText().replace("：", ""),
                "元素值": self.lineEdit_10.text(),
                "工作表": self.comboBox_23.currentText(),
                "单元格": self.lineEdit_9.text(),
                "行号递增": str(self.checkBox_6.isChecked()),
                "超时类型": parameter_4_,
            }
            return image_path_, parameter_dic_

        def put_parameters(image_path_, parameter_dic_):
            """Restore parameters to the widget"""
            # Split the image path into two parts and set the comboBox texts
            self.comboBox_20.setCurrentIndex(self.comboBox_20.findText(image_path_))
            self.find_controls("excel", "网页录入")
            self.comboBox_23.setCurrentIndex(
                self.comboBox_23.findText(parameter_dic_["工作表"])
            )
            # Set the text of the comboBox and lineEdits
            self.comboBox_24.setCurrentText(parameter_dic_["元素类型"] + "：")
            self.lineEdit_10.setText(parameter_dic_["元素值"])
            self.lineEdit_9.setText(parameter_dic_["单元格"])
            # Set the checked state of the checkBox
            self.checkBox_6.setChecked(parameter_dic_["行号递增"] == "True")
            # Set the checked state of the radioButtons and the value of the spinBox
            if parameter_dic_["超时类型"] == "自动跳过":
                self.radioButton_10.setChecked(True)
                self.radioButton_11.setChecked(False)
            else:
                self.radioButton_10.setChecked(False)
                self.radioButton_11.setChecked(True)
                self.spinBox_8.setValue(int(parameter_dic_["超时类型"]))

        if type_ == "按钮功能":
            # 网页信息录入的excel功能
            self.comboBox_20.activated.connect(
                lambda: self.find_controls("excel", "网页录入")
            )
        elif type_ == "写入参数":
            image_path, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                image_=image_path,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 加载文件路径
            self.comboBox_20.clear()
            self.comboBox_20.addItems(extract_excel_from_global_parameter())
            self.comboBox_23.clear()
            self.find_controls("excel", "网页录入")

        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def mouse_drag_function(self, type_):
        """鼠标拖拽窗口的功能"""

        def test():
            """测试功能"""
            try:
                parameter_1_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_1_,
                )
                # 测试用例
                try:
                    mouse_drag = MouseDrag(self.out_mes, dic_)
                    mouse_drag.is_test = True
                    mouse_drag.start_execute()
                except Exception as e:
                    print(e)
                    self.out_mes.out_mes(f"参数错误请重试，测试结束", True)

            except FileNotFoundError:
                self.out_mes.out_mes(f"图像文件未设置！", True)

        def get_parameters():
            """获取参数"""
            parameter_dic_ = {
                "开始位置": f"{self.label_59.text()},{self.label_61.text()}",
                "结束位置": f"{self.label_65.text()},{self.label_66.text()}",
                "开始随机": str(self.checkBox_8.isChecked()),
                "结束随机": str(self.checkBox_7.isChecked()),
                "移动速度": self.spinBox_32.value(),
            }
            if self.label_59.text() == "0" and self.label_61.text() == "0":
                QMessageBox.critical(self, "错误", "未设置开始位置！")
                raise ValueError
            if self.label_65.text() == "0" and self.label_66.text() == "0":
                QMessageBox.critical(self, "错误", "未设置结束位置！")
                raise ValueError
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            # 还原开始位置
            x, y = parameter_dic_["开始位置"].split(",")
            self.label_59.setText(x)
            self.label_61.setText(y)
            # 还原结束位置
            x, y = parameter_dic_["结束位置"].split(",")
            self.label_65.setText(x)
            self.label_66.setText(y)
            # 还原移动速度
            self.spinBox_32.setValue(int(parameter_dic_["移动速度"]))
            # 还原开始随机
            self.checkBox_8.setChecked(parameter_dic_["开始随机"] == "True")
            # 还原结束随机
            self.checkBox_7.setChecked(parameter_dic_["结束随机"] == "True")

        def method_one():
            """方法一"""
            # 获取“运行Python”标题的索引
            tab_index = self.tab_title_list.index('运行Python')
            self.tabWidget.setCurrentIndex(tab_index)
            code_1 = (
                "import pyautogui\n\n"
                "var_1 =  eval( ) # 括号里插入开始位置的变量\n"
                "var_2 =  eval( ) # 括号里插入结束位置的变量\n"
                "duration_time = 0.3  # 此处填写移动时间（单位：秒s）\n\n"
                "pyautogui.moveTo(var_1[0], var_1[1], duration=duration_time)\n"
                "pyautogui.dragTo(var_2[0], var_2[1], duration=duration_time)"
            )
            self.textEdit_5.setText(code_1)

        if type_ == "按钮功能":
            # 鼠标拖拽
            self.pushButton_12.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "开始拖拽"
                )
            )
            # self.pushButton_12.clicked.connect(self.mouseMoveEvent)

            self.pushButton_13.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "结束拖拽"
                )
            )
            # self.pushButton_13.clicked.connect(self.mouseMoveEvent)
            # 拖拽测试按钮
            self.pushButton_14.clicked.connect(test)
            self.pushButton_83.clicked.connect(method_one)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def toggle_frame_function(self, type_):
        """切换frame窗口的功能"""

        def switch_frame():
            """切换frame"""
            # 切换frame时控件的状态
            if self.comboBox_26.currentText() == "切换到指定frame":
                self.comboBox_27.setVisible(True)
                self.lineEdit_11.clear()
                self.lineEdit_11.setVisible(True)
            else:
                self.comboBox_27.setVisible(False)
                self.lineEdit_11.clear()
                self.lineEdit_11.setVisible(False)

        def get_parameters():
            """获取参数"""
            # 检查参数是否有异常
            if self.comboBox_26.currentText() == "切换到指定frame" and not self.lineEdit_11.text():
                QMessageBox.critical(self, "错误", "未设置frame！")
                raise ValueError
            # 获取参数字典
            if self.comboBox_26.currentText() == "切换到指定frame":
                parameter_dic_ = {
                    "指令类型": self.comboBox_26.currentText(),
                    "frame类型": self.comboBox_27.currentText().replace("：", ""),
                    "frame": self.lineEdit_11.text()
                }
                return parameter_dic_
            else:
                parameter_dic_ = {
                    "指令类型": self.comboBox_26.currentText(),
                }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            # 还原指令类型
            self.comboBox_26.setCurrentText(parameter_dic_["指令类型"])
            switch_frame()
            # 还原frame类型
            if parameter_dic_["指令类型"] == "切换到指定frame":
                self.comboBox_27.setCurrentText(parameter_dic_["frame类型"])
                self.lineEdit_11.setText(parameter_dic_["frame"])

        if type_ == "按钮功能":
            # 切换frame
            self.comboBox_26.activated.connect(switch_frame)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def save_form_function(self, type_):
        """保存网页表格的功能"""

        def get_parameters():
            """获取参数"""
            # 检查参数是否有异常
            if self.lineEdit_12.text() == "":
                QMessageBox.critical(self, "错误", "元素未填写！")
                raise ValueError
            if self.comboBox_29.currentText() == "":
                QMessageBox.critical(self, "错误", "未设置工作簿！")
                raise ValueError
            if self.lineEdit_13.text() == "":
                QMessageBox.critical(self, "错误", "未填写工作表名！")
                raise ValueError
            # 异常处理
            timeout_type = "自动跳过" if self.radioButton_13.isChecked() \
                else self.spinBox_9.value()
            # 获取参数字典
            image_ = self.lineEdit_12.text()  # 元素
            parameter_dic_ = {
                "工作簿": self.comboBox_29.currentText(),
                "工作表": self.lineEdit_13.text(),
                "元素类型": self.comboBox_28.currentText().replace("：", ""),
                "异常": timeout_type,
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原元素
            self.lineEdit_12.setText(image_)
            # 还原工作簿
            self.comboBox_29.setCurrentText(parameter_dic_["工作簿"])
            # 还原工作表
            self.lineEdit_13.setText(parameter_dic_["工作表"])
            # 还原元素类型
            self.comboBox_28.setCurrentText(parameter_dic_["元素类型"] + "：")
            # 还原异常处理
            if parameter_dic_["异常"] == "自动跳过":
                self.radioButton_13.setChecked(True)
                self.radioButton_12.setChecked(False)
            else:
                self.radioButton_13.setChecked(False)
                self.radioButton_12.setChecked(True)
                self.spinBox_9.setValue(int(parameter_dic_["异常"]))

        if type_ == "按钮功能":
            pass
        elif type_ == "写入参数":
            # 获取参数
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                image_=image,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            self.comboBox_29.clear()
            self.comboBox_29.addItems(extract_excel_from_global_parameter())
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def drag_element_function(self, type_):
        """拖动网页元素的功能"""

        def get_parameters():
            """获取参数"""
            # 检查参数是否有异常
            if self.lineEdit_14.text() == "":
                QMessageBox.critical(self, "错误", "元素未填写！")
                raise ValueError
            if self.spinBox_10.value() == 0 and self.spinBox_11.value() == 0:
                QMessageBox.critical(self, "错误", "未设置拖动距离！")
                raise ValueError
            # 获取参数字典
            image_ = self.lineEdit_14.text()
            parameter_dic_ = {
                "距离X": self.spinBox_10.value(),
                "距离Y": self.spinBox_11.value(),
                "异常": "自动跳过" if self.radioButton_15.isChecked() else self.spinBox_12.value(),
                "元素类型": self.comboBox_30.currentText().replace("：", ""),
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原元素
            self.lineEdit_14.setText(image_)
            # 还原拖动距离
            x = parameter_dic_["距离X"]
            y = parameter_dic_["距离Y"]
            self.spinBox_10.setValue(int(x))
            self.spinBox_11.setValue(int(y))
            # 还原元素类型
            self.comboBox_30.setCurrentText(parameter_dic_["元素类型"] + "：")
            # 还原异常处理
            if parameter_dic_["异常"] == "自动跳过":
                self.radioButton_15.setChecked(True)
                self.radioButton_14.setChecked(False)
            else:
                self.radioButton_15.setChecked(False)
                self.radioButton_14.setChecked(True)
                self.spinBox_12.setValue(int(parameter_dic_["异常"]))

        if type_ == "按钮功能":
            pass
        elif type_ == "写入参数":
            # 获取参数
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                image_=image,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def full_screen_capture_function(self, type_):
        """全屏截图的窗口功能"""

        def show_region():
            """显示区域"""
            is_region_screenshot = self.comboBox_67.currentText() == "区域截图"
            self.label_164.setVisible(is_region_screenshot)
            self.pushButton_60.setVisible(is_region_screenshot)

        def show_save_path():
            """显示保存路径"""
            is_save_path = self.radioButton_9.isChecked()
            self.groupBox_14.setVisible(is_save_path)

        def get_parameters():
            """获取参数"""
            # 检查参数是否有异常
            if self.comboBox_67.currentText() == "区域截图" and self.label_164.text() == "(0,0,0,0)":
                QMessageBox.critical(self, "错误", "未设置区域！")
                raise ValueError
            if self.radioButton_9.isChecked() and self.lineEdit_16.text() == "":
                QMessageBox.critical(self, "错误", "未设置图像名称！")
                raise ValueError
            # 获取参数字典
            if not self.lineEdit_16.text().endswith(".png"):  # 如果没有.png后缀则添加
                self.lineEdit_16.setText(self.lineEdit_16.text() + ".png")
            image_path_ = os.path.join(self.comboBox_31.currentText(), self.lineEdit_16.text())
            parameter_dic_ = {
                "截图类型": self.comboBox_67.currentText(),
                "区域": self.label_164.text(),
                "截图后": "保存到路径" if self.radioButton_9.isChecked() else "写入剪切板",
            }
            return image_path_, parameter_dic_

        def put_parameters(image_path_, parameter_dic_):
            """将参数还原到控件中"""
            # 还原截图类型
            self.comboBox_67.setCurrentText(parameter_dic_["截图类型"])
            show_region()
            # 还原区域
            self.label_164.setText(parameter_dic_["区域"])
            # 还原截图后
            if parameter_dic_["截图后"] == "保存到路径":
                self.radioButton_9.setChecked(True)
                self.radioButton_8.setChecked(False)
                show_save_path()
            else:
                self.radioButton_9.setChecked(False)
                self.radioButton_8.setChecked(True)
                show_save_path()
            # 还原图像路径
            self.comboBox_31.setCurrentText(os.path.split(image_path_)[0])
            self.lineEdit_16.setText(os.path.split(image_path_)[1])

        def test():
            """测试功能"""
            try:
                image_path_, parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                    image_=image_path_,
                )
                # 测试用例
                test_class = FullScreenCapture(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            self.pushButton_60.clicked.connect(
                lambda: self.quick_screenshot(self.label_164, "设置区域")
            )
            self.comboBox_67.activated.connect(show_region)
            self.radioButton_9.clicked.connect(show_save_path)
            self.radioButton_8.clicked.connect(show_save_path)
            # 测试按钮
            self.pushButton_61.clicked.connect(test)
            # 打开文件夹
            self.pushButton_62.clicked.connect(
                lambda: self.quick_screenshot(self.comboBox_31, "打开文件夹")
            )
        elif type_ == "写入参数":
            # 获取参数
            image_path, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image_path,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            self.comboBox_31.clear()
            self.comboBox_31.addItems(extract_resource_folder_path())
            show_region()
            show_save_path()
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def switch_window_function(self, type_):
        """切换浏览器窗口的功能"""

        def get_parameters():
            """获取参数"""
            # 检查参数是否有异常
            if self.lineEdit_15.text() == "":
                QMessageBox.critical(self, "错误", "窗口未填写！")
                raise ValueError
            # 获取参数字典
            parameter_dic_ = {
                "窗口": self.lineEdit_15.text(),
                "窗口类型": self.comboBox_32.currentText().replace("：", ""),
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件中"""
            # 还原窗口
            self.lineEdit_15.setText(parameter_dic_["窗口"])
            # 还原窗口类型
            self.comboBox_32.setCurrentText(parameter_dic_["窗口类型"] + "：")

        if type_ == "按钮功能":
            pass
        elif type_ == "写入参数":
            # 获取参数
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def wechat_function(self, type_):
        """微信发送消息的功能"""

        def Lock_control():
            """锁定控件"""
            if self.comboBox_33.currentText() == "自定义联系人":
                self.lineEdit_17.setEnabled(True)
            else:
                self.lineEdit_17.setEnabled(False)
                self.lineEdit_17.clear()

            if self.comboBox_34.currentText() == "自定义消息内容":
                self.textEdit_2.setEnabled(True)
            else:
                self.textEdit_2.setEnabled(False)
                self.textEdit_2.clear()

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = (
                self.comboBox_33.currentText()
                if self.comboBox_33.currentText() == "文件传输助手"
                else self.lineEdit_17.text()
            )
            parameter_2_ = (
                self.comboBox_34.currentText()
                if self.comboBox_34.currentText() != "自定义消息内容"
                else self.textEdit_2.toPlainText()
            )
            if parameter_1_ == "" or parameter_2_ == "":
                QMessageBox.critical(self, "错误", "联系人或消息内容不能为空！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "联系人": parameter_1_,
                "消息内容": parameter_2_,
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            contact = parameter_dic_.get("联系人", "")
            message = parameter_dic_.get("消息内容", "")
            print(contact, message)
            # 设置联系人
            if contact == "文件传输助手":
                self.comboBox_33.setCurrentText(contact)
                self.lineEdit_17.setEnabled(False)
            else:
                self.comboBox_33.setCurrentText("自定义联系人")
                self.lineEdit_17.setEnabled(True)
                self.lineEdit_17.setText(contact)
            # 设置消息内容
            if message in ['从剪切板粘贴', '当前日期时间']:
                self.comboBox_34.setCurrentText(message)
                self.textEdit_2.setEnabled(False)
            else:
                self.comboBox_34.setCurrentText("自定义消息内容")
                self.textEdit_2.setEnabled(True)
                self.textEdit_2.setText(message)

        def test():
            """测试"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )
                # 测试用例
                test_class = SendWeChat(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            Lock_control()
            self.comboBox_33.activated.connect(Lock_control)
            self.comboBox_34.activated.connect(Lock_control)
            self.pushButton_15.clicked.connect(test)
            self.pushButton_30.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )

        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def verification_code_function(self, type_):
        """数字验证码功能"""

        def test():
            """测试功能"""
            try:
                parameter_1_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()), parameter_1_=parameter_1_
                )
                # 测试用例
                verification_code = VerificationCode(self.out_mes, dic_)
                verification_code.is_test = True
                verification_code.start_execute()
            except Exception as e:
                self.out_mes.out_mes(f"识别失败，错误信息：{type(e)}", True)

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

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = self.label_85.text()  # 截图区域
            parameter_2_ = self.comboBox_63.currentText()  # 变量
            parameter_3_ = self.comboBox_62.currentText()  # 验证码类型
            # 检查参数是否有异常
            if parameter_1_ == "(0,0,0,0)":
                QMessageBox.critical(self, "错误", "验证码识别区域未设置！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "区域": parameter_1_,
                "变量": parameter_2_,
                "验证码类型": parameter_3_,
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到tab页"""
            # 设置截图区域
            self.label_85.setText(parameter_dic_["区域"])
            # 设置变量
            index = self.comboBox_63.findText(parameter_dic_["变量"])
            if index >= 0:
                self.comboBox_63.setCurrentIndex(index)
            # 设置验证码类型
            index = self.comboBox_62.findText(parameter_dic_["验证码类型"])
            if index >= 0:
                self.comboBox_62.setCurrentIndex(index)

        if type_ == "按钮功能":
            self.pushButton_16.clicked.connect(set_region)
            self.pushButton_53.clicked.connect(open_setting_window)
            self.pushButton_55.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            # 测试按钮
            self.pushButton_17.clicked.connect(test)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )

        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

        elif type_ == "加载信息":
            self.comboBox_63.clear()
            self.comboBox_63.addItems(get_variable_info("list"))

    def play_voice_function(self, type_):
        """播放语音的功能"""

        def get_parameters():
            # 检查参数是否有异常
            if self.groupBox_34.isChecked() and not self.textEdit_4.toPlainText():
                QMessageBox.critical(self, "错误", "内容未输入！")
                raise ValueError
            if self.groupBox_32.isChecked():
                parameter_dic_ = {
                    "类型": "音频信号",
                    "频率": self.spinBox_21.value(),
                    "持续": self.spinBox_23.value(),
                    "次数": self.spinBox_22.value(),
                    "间隔": self.spinBox_24.value(),
                }
                return parameter_dic_
            elif self.groupBox_33.isChecked():
                parameter_dic_ = {
                    "类型": "系统提示音",
                    "提示类型": self.comboBox_7.currentText(),
                }
                return parameter_dic_
            elif self.groupBox_34.isChecked():
                parameter_dic_ = {
                    "类型": "播放语音",
                    "内容": self.textEdit_4.toPlainText(),
                    "语速": self.horizontalSlider.value(),
                }
                return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            all_groupBoxes__ = [self.groupBox_32, self.groupBox_33, self.groupBox_34]
            if parameter_dic_["类型"] == "音频信号":
                self.groupBox_32.setChecked(True)
                self.select_groupBox(self.groupBox_32, all_groupBoxes__)
                self.spinBox_21.setValue(int(parameter_dic_["频率"]))
                self.spinBox_23.setValue(int(parameter_dic_["持续"]))
                self.spinBox_22.setValue(int(parameter_dic_["次数"]))
                self.spinBox_24.setValue(int(parameter_dic_["间隔"]))
            elif parameter_dic_["类型"] == "系统提示音":
                self.groupBox_33.setChecked(True)
                self.select_groupBox(self.groupBox_33, all_groupBoxes__)
                index = self.comboBox_7.findText(parameter_dic_["提示类型"])
                if index >= 0:
                    self.comboBox_7.setCurrentIndex(index)
            elif parameter_dic_["类型"] == "播放语音":
                self.groupBox_34.setChecked(True)
                self.select_groupBox(self.groupBox_34, all_groupBoxes__)
                self.textEdit_4.setText(parameter_dic_["内容"])
                self.horizontalSlider.setValue(int(parameter_dic_["语速"]))
                self.label_118.setText(str(parameter_dic_["语速"]))

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )
                play_voice = PlayVoice(self.out_mes, dic_)
                play_voice.is_test = True
                play_voice.start_execute()
            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"参数异常", True)

        if type_ == "按钮功能":
            all_groupBoxes_ = [self.groupBox_32, self.groupBox_33, self.groupBox_34]
            for groupBox_ in all_groupBoxes_:
                groupBox_.clicked.connect(lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_))
            # 测试按钮
            self.pushButton_24.clicked.connect(test)
            self.pushButton_32.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            pass
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def wait_window_function(self, type_):
        """倒计时等待窗口的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_1_ = f"{self.lineEdit_2.text()}" or "示例"
            parameter_2_ = f"{self.lineEdit_6.text()}" or "示例"
            parameter_3_ = f"{self.spinBox_25.value()}"
            # 返回参数字典
            parameter_dic_ = {
                "标题": parameter_1_,
                "内容": parameter_2_,
                "秒数": parameter_3_,
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.lineEdit_2.setText(parameter_dic_["标题"])
            self.lineEdit_6.setText(parameter_dic_["内容"])
            self.spinBox_25.setValue(int(parameter_dic_["秒数"]))

        def test():
            # """测试功能"""
            parameter_dic_ = get_parameters()
            dic_ = self.get_test_dic(
                repeat_number_=int(self.spinBox.value()),
                parameter_1_=parameter_dic_,
            )
            # 测试用例
            test_class = WaitWindow(self.out_mes, dic_)
            test_class.is_test = True
            test_class.start_execute()

        if type_ == "按钮功能":
            self.pushButton_25.clicked.connect(test)
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            pass
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def dialog_window_function(self, type_):
        """弹出提示框的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            title = self.lineEdit_8.text() or "提示框"  # 提示框标题
            info = self.lineEdit_20.text() or "示例"  # 提示框内容
            icon_type = self.comboBox_36.currentText()  # icon类型
            # 返回参数字典
            parameter_dic_ = {
                "标题": title,
                "内容": info,
                "图标": icon_type,
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.lineEdit_8.setText(parameter_dic_["标题"])
            self.lineEdit_20.setText(parameter_dic_["内容"])
            index = self.comboBox_36.findText(parameter_dic_["图标"])
            if index >= 0:
                self.comboBox_36.setCurrentIndex(index)

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    parameter_1_=parameter_dic_,
                    repeat_number_=int(self.spinBox.value()),
                )
                # 测试用例
                test_class = DialogWindow(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误！", True)

        if type_ == "按钮功能":
            self.pushButton_26.clicked.connect(test)

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            pass
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def branch_jump_function(self, type_):
        """跳转分支的功能
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.comboBox_37.currentText() == "" or self.comboBox_38.currentText() == "":
                QMessageBox.critical(self, "错误", "分支参数错误！")
                raise ValueError
            # 返回参数字典
            parameter_dic = {
                "分支": f"{self.comboBox_37.currentText()}-{self.comboBox_38.currentText()}"
            }
            exception_handling = f"{self.comboBox_37.currentText()}-{self.comboBox_38.currentText()}"
            return exception_handling, parameter_dic

        def put_parameters(parameter_dic):
            """将参数还原到控件"""
            # 设置分支
            branch = parameter_dic["分支"]
            branch_name, branch_count = branch.split("-")
            index = self.comboBox_37.findText(branch_name)
            if index >= 0:
                self.comboBox_37.setCurrentIndex(index)
            # 设置跳转分支
            index = self.comboBox_38.findText(branch_count)
            if index >= 0:
                self.comboBox_38.setCurrentIndex(index)

        if type_ == "按钮功能":
            self.comboBox_37.activated.connect(
                lambda: self.find_controls("分支", "跳转分支")
            )

        elif type_ == "写入参数":
            exception_handling_, parameter_dic_ = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=exception_handling_,
                parameter_1_=parameter_dic_,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            self.comboBox_37.addItems(get_branch_info(True))
            self.comboBox_37.setCurrentIndex(0)
            # 获取分支表名中的指令数量
            self.find_controls("分支", "跳转分支")
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def termination_process_function(self, type_):
        """终止流程的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            exception_handling = self.comboBox_39.currentText()
            # 返回参数字典
            parameter_dic = {
                "终止类型": exception_handling,
            }
            return exception_handling, parameter_dic

        def put_parameters(parameter_dic):
            """将参数还原到控件"""
            # 设置终止类型
            index = self.comboBox_39.findText(parameter_dic["终止类型"])
            if index >= 0:
                self.comboBox_39.setCurrentIndex(index)

        if type_ == "按钮功能":
            pass

        elif type_ == "写入参数":
            exception_handling_, parameter_dic_ = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=exception_handling_,
                parameter_1_=parameter_dic_,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def window_control_function(self, type_):
        """窗口控制的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.lineEdit_21.text() == "":
                QMessageBox.critical(self, "错误", "窗口标题未填！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "标题包含": self.lineEdit_21.text(),
                "操作": self.comboBox_40.currentText(),
                "报错": str(self.checkBox_5.isChecked()),
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.lineEdit_21.setText(parameter_dic_["标题包含"])
            index = self.comboBox_40.findText(parameter_dic_["操作"])
            if index >= 0:
                self.comboBox_40.setCurrentIndex(index)
            self.checkBox_5.setChecked(eval(parameter_dic_["报错"]))

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )
                # 测试用例
                test_class = WindowControl(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == "按钮功能":
            self.pushButton_27.clicked.connect(test)

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def key_wait_function(self, type_):
        """按键等待的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            key_name = self.keySequenceEdit_2.keySequence().toString()
            if key_name == "":
                QMessageBox.critical(self, "错误", "按键未设置！")
                raise ValueError
            if key_name.count("+") >= 1:
                QMessageBox.critical(self, "错误", "该功能暂不支持复合按键！")
                raise ValueError
            if self.radioButton_21.isChecked() and (
                    self.comboBox_41.currentText() == ""
                    or self.comboBox_42.currentText() == ""
            ):
                QMessageBox.critical(self, "错误", "分支异常，请先添加！")
                raise ValueError
            # 返回参数字典
            if self.radioButton_22.isChecked():  # 按键等待
                parameter_dic_ = {
                    "按键": key_name,
                    "等待类型": "按键等待",
                }
                exception_handling = '提示异常并暂停'
                return parameter_dic_, exception_handling
            elif self.radioButton_21.isChecked():  # 跳转分支
                parameter_dic_ = {
                    "按键": key_name,
                    "等待类型": "跳转分支",
                    "分支": f"{self.comboBox_41.currentText()}-{self.comboBox_42.currentText()}",
                }
                exception_handling = f"{self.comboBox_41.currentText()}-{self.comboBox_42.currentText()}"
                return parameter_dic_, exception_handling

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            # 设置按键
            self.keySequenceEdit_2.setKeySequence(QKeySequence(parameter_dic_["按键"]))
            # 设置等待类型
            if parameter_dic_["等待类型"] == "按键等待":
                self.radioButton_22.setChecked(True)
            elif parameter_dic_["等待类型"] == "跳转分支":
                self.radioButton_21.setChecked(True)
                # 设置分支
                self.comboBox_41.setCurrentText(parameter_dic_["分支"].split("-")[0])
                self.comboBox_42.setCurrentText(parameter_dic_["分支"].split("-")[1])

        def set_branch_name():
            """当选择跳转分支功能时，加载分支表名"""
            disable_control(True)
            self.comboBox_41.addItems(get_branch_info(True))
            self.find_controls("分支", "按键等待")

        def disable_control(judge_: bool):
            """禁用控件"""
            self.comboBox_41.clear()
            self.comboBox_42.clear()
            self.label_133.setEnabled(judge_)
            self.label_132.setEnabled(judge_)
            self.comboBox_41.setEnabled(judge_)
            self.comboBox_42.setEnabled(judge_)

        if type_ == "按钮功能":
            self.radioButton_21.toggled.connect(set_branch_name)
            self.radioButton_22.toggled.connect(lambda: disable_control(False))
            self.comboBox_41.activated.connect(
                lambda: self.find_controls("分支", "按键等待")
            )

        elif type_ == "写入参数":
            parameter_dic, exception_handling_ = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=exception_handling_,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def gain_time_function(self, type_):
        """获取时间的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.comboBox_44.currentText() == "":
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "变量": self.comboBox_44.currentText(),
                "时间格式": self.comboBox_43.currentText(),
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            # 设置变量
            index = self.comboBox_44.findText(parameter_dic_["变量"])
            if index >= 0:
                self.comboBox_44.setCurrentIndex(index)
            # 设置时间格式
            index = self.comboBox_43.findText(parameter_dic_["时间格式"])
            if index >= 0:
                self.comboBox_43.setCurrentIndex(index)

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )

                # 测试用例
                test_class = GetTimeValue(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            self.pushButton_33.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.pushButton_34.clicked.connect(test)

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_44.clear()
            self.comboBox_44.addItems(get_variable_info("list"))
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def gain_excel_function(self, type_):
        """从excel单元格中获取变量的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if (
                    self.comboBox_45.currentText() == ""
                    or self.comboBox_46.currentText() == ""
            ):
                QMessageBox.critical(self, "错误", "Excel路径未设置！")
                raise ValueError
            if self.lineEdit_23.text() == "":
                QMessageBox.critical(self, "错误", "单元格未设置！")
                raise ValueError
            if self.comboBox_47.currentText() == "":
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            # 返回参数字典
            image_ = f"{self.comboBox_45.currentText()}"
            parameter_dic_ = {
                "工作表": self.comboBox_46.currentText(),
                "单元格": self.lineEdit_23.text(),
                "变量": self.comboBox_47.currentText(),
                "递增": str(self.checkBox_9.isChecked()),
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到控件"""
            self.comboBox_45.setCurrentText(image_)
            self.comboBox_46.setCurrentText(parameter_dic_["工作表"])
            self.lineEdit_23.setText(parameter_dic_["单元格"])
            self.comboBox_47.setCurrentText(parameter_dic_["变量"])
            self.checkBox_9.setChecked(eval(parameter_dic_["递增"]))

        def test():
            """测试功能"""
            try:
                image_, parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    image_=image_,
                    parameter_1_=parameter_dic_,
                )

                # 测试用例
                test_class = GetExcelCellValue(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        def line_number_increasing():
            # 行号递增功能被选中后弹出提示框
            if self.checkBox_9.isChecked():
                QMessageBox.information(
                    self,
                    "提示",
                    "启用该功能后，请在主页面中设置循环次数大于1，执行全部指令后，"
                    "循环执行时，单元格行号会自动递增。",
                    QMessageBox.Ok,
                )

        if type_ == "按钮功能":
            # 禁用中文输入
            self.lineEdit_23.setValidator(
                QRegExpValidator(QRegExp("[a-zA-Z0-9]{16}"), self)
            )
            self.checkBox_9.clicked.connect(line_number_increasing)
            self.comboBox_45.activated.connect(
                lambda: self.find_controls("excel", "获取Excel")
            )
            self.pushButton_35.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.pushButton_36.clicked.connect(test)
            # 打开工作簿
            self.pushButton_29.clicked.connect(
                lambda: os.startfile(self.comboBox_45.currentText())
            )

        elif type_ == "写入参数":
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_45.clear()
            self.comboBox_45.addItems(
                extract_excel_from_global_parameter()
            )  # 加载全局参数中的excel文件路径
            self.find_controls("excel", "获取Excel")

            self.comboBox_47.clear()
            self.comboBox_47.addItems(get_variable_info("list"))
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def get_dialog_function(self, type_):
        """从对话框中获取变量的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.comboBox_48.currentText() == "":
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            # 参数字典
            parameter_dic_ = {
                "标题": self.lineEdit_24.text() if self.lineEdit_24.text() else "示例",
                "变量": self.comboBox_48.currentText(),
                "提示": self.lineEdit_25.text() if self.lineEdit_25.text() else "示例",
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.lineEdit_24.setText(parameter_dic_.get("标题", ""))
            if self.lineEdit_24.text() == "示例":
                self.lineEdit_24.clear()
            self.comboBox_48.setCurrentText(parameter_dic_.get("变量", ""))
            self.lineEdit_25.setText(parameter_dic_.get("提示", ""))
            if self.lineEdit_25.text() == "示例":
                self.lineEdit_25.clear()

        if type_ == "按钮功能":
            self.pushButton_37.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_48.clear()
            self.comboBox_48.addItems(get_variable_info("list"))
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def get_clipboard_function(self, type_):
        """从剪切板中获取变量的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.comboBox_73.currentText() == "":
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                '变量': self.comboBox_73.currentText(),
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到tab页"""
            index = self.comboBox_73.findText(parameter_dic_['变量'])
            if index >= 0:
                self.comboBox_73.setCurrentIndex(index)

        def test():
            """测试功能"""
            try:
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_
                )
                # 测试用例
                test_class = GetClipboard(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_74.clicked.connect(
                lambda: self.merge_additional_functions('打开变量池')
            )
            self.pushButton_75.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=func_info_dic.get('异常处理'),
                                                 parameter_1_=parameter_dic,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            self.comboBox_73.clear()
            self.comboBox_73.addItems(get_variable_info("list"))
        elif type_ == '还原参数':
            put_parameters(self.parameter_1)

    def contrast_variables_function(self, type_):
        """变量比较的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_dic_ = {
                "变量1": self.comboBox_49.currentText(),
                "类型1": self.comboBox_54.currentText(),
                "比较符": self.comboBox_50.currentText(),
                "变量2": self.comboBox_51.currentText(),
                "类型2": self.comboBox_55.currentText(),
                "分支": self.comboBox_52.currentText(),
                "位置": self.comboBox_53.currentText(),
            }
            # 比较符-变量类型
            exception_handling_ = (
                f"{self.comboBox_52.currentText()}" f"-{self.comboBox_53.currentText()}"
            )  # 分支表名-分支序号
            # 检查参数是否有异常
            if (
                    self.comboBox_49.currentText() == ""
                    or self.comboBox_50.currentText() == ""
                    or self.comboBox_51.currentText() == ""
            ):
                QMessageBox.critical(self, "错误", "变量未设置！")
                raise ValueError
            if (
                    self.comboBox_52.currentText() == ""
                    or self.comboBox_53.currentText() == ""
            ):
                QMessageBox.critical(self, "错误", "分支未设置！")
                raise ValueError
            return parameter_dic_, exception_handling_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.comboBox_49.setCurrentText(parameter_dic_.get("变量1", ""))
            self.comboBox_54.setCurrentText(parameter_dic_.get("类型1", ""))
            self.comboBox_50.setCurrentText(parameter_dic_.get("比较符", ""))
            self.comboBox_51.setCurrentText(parameter_dic_.get("变量2", ""))
            self.comboBox_55.setCurrentText(parameter_dic_.get("类型2", ""))
            self.comboBox_52.setCurrentText(parameter_dic_.get("分支", ""))
            self.comboBox_53.setCurrentText(parameter_dic_.get("位置", ""))

        def sync_combo_boxes(sender):
            if sender == self.comboBox_54:
                self.comboBox_55.setCurrentIndex(self.comboBox_54.currentIndex())
            else:
                self.comboBox_54.setCurrentIndex(self.comboBox_55.currentIndex())

        if type_ == "按钮功能":
            self.comboBox_52.activated.connect(  # 当分支表名改变时，加载分支中的命令序号
                lambda: self.find_controls("分支", "变量判断")
            )
            self.pushButton_38.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.comboBox_54.currentIndexChanged.connect(
                lambda: sync_combo_boxes(self.comboBox_54)
            )
            self.comboBox_55.currentIndexChanged.connect(
                lambda: sync_combo_boxes(self.comboBox_55)
            )

        elif type_ == "写入参数":
            parameter_dic, exception_handling = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=exception_handling,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_49.clear()
            self.comboBox_49.addItems(get_variable_info("list"))
            self.comboBox_51.clear()
            self.comboBox_51.addItems(get_variable_info("list"))
            self.comboBox_52.clear()
            self.comboBox_52.addItems(get_branch_info(True))
            self.comboBox_52.setCurrentIndex(0)
            # 获取分支表名中的指令数量
            self.find_controls("分支", "变量判断")
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def run_python_function(self, type_):
        """运行python代码的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = self.textEdit_5.toPlainText()  # 代码
            # 检查参数是否有异常
            if image_ == "":
                QMessageBox.critical(self, "错误", "代码未编写！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "返回值": self.lineEdit_26.text(),
                "变量": self.comboBox_56.currentText(),
            }
            return image_, parameter_dic_

        def put_parameters(image_, parameter_dic_):
            """将参数还原到tab页"""
            self.textEdit_5.setPlainText(image_)
            self.lineEdit_26.setText(parameter_dic_["返回值"])
            index = self.comboBox_56.findText(parameter_dic_["变量"])
            if index >= 0:
                self.comboBox_56.setCurrentIndex(index)

        def test():
            """测试功能"""
            highlight_python_code()
            try:
                image_, parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    image_=image_,
                    parameter_1_=parameter_dic_,
                )
                # 测试用例
                test_class = RunPython(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        def highlight_python_code():
            """运行python代码"""

            def highlight_text(text):
                lexer = PythonLexer()
                formatter = HtmlFormatter(style="monokai")
                html = highlight(text, lexer, formatter)
                css = formatter.get_style_defs(".highlight")
                self.textEdit_5.setHtml("<style>" + css + "</style>" + html)

            code = self.textEdit_5.toPlainText()
            highlight_text(code)

        def show_lib_info():
            """显示库的信息"""
            title = ["模块名称", "说明"]
            data = [
                ("pyttsx4", "文本转语音"),
                ("pymsgbox", "消息框"),
                ("pyautogui", "自动化GUI，鼠标、键盘控制"),
                ("mouse", "鼠标控制"),
                ("keyboard", "键盘控制"),
                ("pandas", "数据处理"),
                ("selenium", "网页自动化"),
                ("pillow", "图像处理"),
                ("openpyxl", "Excel操作"),
                ("requests", "HTTP请求"),
                ("python-dateutil", "日期处理"),
                ("psutil", "系统监控"),
                ("pywinauto", "Windows自动化")
            ]
            shortcut_win = ShortcutTable(self, title, data,600)  # 快捷键说明窗口
            shortcut_win.setWindowTitle("库的使用")
            shortcut_win.setModal(True)
            shortcut_win.exec_()

        if type_ == "按钮功能":
            # 自动代码高亮
            self.pushButton_40.clicked.connect(test)
            self.pushButton_39.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )
            self.pushButton_41.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.toolButton_4.clicked.connect(show_lib_info)

        elif type_ == "写入参数":
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_56.clear()
            self.comboBox_56.addItems(get_variable_info("list"))
            # 设置textEdit_5的说明信息
            self.textEdit_5.setPlaceholderText(
                "执行python代码......"
                "\n\n已内置的第三方库："
                "\npyttsx4、pymsgbox、pyautogui、mouse、keyboard、pandas、selenium、"
                "pillow、openpyxl、requests、python-dateutil、psutil、pywinauto"
                "\n\n点击帮助按钮查看库的使用"
                "\n\n请去除代码中的"
                "\nif __name__ == '__main__': "
                "\n否则无法执行"
            )
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def run_cmd_function(self, type_):
        """运行cmd命令的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def write_cmd_to_textedit(comboBox_72_text):
            """将命令写入textEdit"""
            cmd_dic = {
                "立即关闭计算机": "shutdown -s -t 0",
                "1分钟后关闭计算机": "shutdown -s -t 60",
                "重启计算机": "shutdown -r -t 0",
                "锁定屏幕": "rundll32.exe user32.dll,LockWorkStation",
                "注销账户": "shutdown -l",
                "创建新目录": "mkdir 目录名",
                "删除目录": "rmdir /s /q 目录名",
                "终止进程": "taskkill 进程名.exe",
                "打开记事本": "notepad",
                "打开计算器": "calc",
                "打开资源管理器": "explorer",
                "打开控制面板": "control",
            }
            self.textEdit_7.setPlainText(
                f"{self.textEdit_7.toPlainText()}\n{cmd_dic[comboBox_72_text]}".strip("\n"))

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.textEdit_7.toPlainText() == "":
                QMessageBox.critical(self, "错误", "命令未填写！")
                raise ValueError
            # 返回参数字典
            image_ = self.textEdit_7.toPlainText()
            return image_

        def put_parameters(image_):
            """将参数还原到tab页"""
            self.textEdit_7.setPlainText(image_)

        def test():
            """测试功能"""
            try:
                image_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    image_=image_,
                )

                # 测试用例
                test_class = RunCmd(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_70.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )
            self.pushButton_71.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.pushButton_73.clicked.connect(
                lambda: write_cmd_to_textedit(self.comboBox_72.currentText())
            )
            self.pushButton_72.clicked.connect(test)

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

        elif type_ == '还原参数':
            put_parameters(self.image_path)

    def run_external_file_function(self, type_):
        """运行外部文件的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            image_ = self.lineEdit_27.text()  # 文件路径
            # 检查参数是否有异常
            if image_ is None or image_ == "":
                QMessageBox.critical(self, "错误", "文件路径未设置！")
                raise ValueError
            return image_

        def test():
            """测试功能"""
            try:
                image_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()), image_=image_
                )

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
                directory=os.path.join(os.path.expanduser("~"), "Desktop"),
            )
            if file_path != "":  # 获取文件名称
                # 设置文件路径
                self.lineEdit_27.setText(os.path.normpath(file_path))

        if type_ == "按钮功能":
            self.pushButton_43.clicked.connect(get_file_and_folder)  # 打开文件选择窗口
            self.pushButton_42.clicked.connect(test)  # 测试按钮

        elif type_ == "写入参数":
            image = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            pass

    def input_cell_function(self, type_):
        """输入到excel单元格的功能
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # excel路径
            image_ = self.comboBox_57.currentText()
            if image_ == "":
                QMessageBox.critical(self, "错误", "Excel路径未设置！")
                raise ValueError
            # 工作表
            if self.comboBox_58.currentText() == "":
                QMessageBox.critical(self, "错误", "工作表未设置！")
                raise ValueError
            if self.lineEdit_28.text() == "":
                QMessageBox.critical(self, "错误", "单元格未设置！")
                raise ValueError
            # 参数字典
            parameter_dic_ = {
                "工作表": self.comboBox_58.currentText(),
                "单元格": self.lineEdit_28.text(),
                "递增": str(self.checkBox_10.isChecked()),
                "文本": self.textEdit_6.toPlainText(),
            }
            return image_, parameter_dic_

        def put_parameters(image, parameter_dic_):
            """将参数还原到控件"""
            self.comboBox_57.setCurrentText(image)
            self.find_controls("excel", "写入单元格")
            self.comboBox_58.setCurrentText(parameter_dic_.get("工作表", ""))
            self.lineEdit_28.setText(parameter_dic_.get("单元格", ""))
            self.checkBox_10.setChecked(eval(parameter_dic_.get("递增", False)))
            self.textEdit_6.setText(parameter_dic_.get("文本", ""))

        def test():
            """测试功能"""
            try:
                image_, parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    image_=image_,
                    parameter_1_=parameter_dic_,
                )
                # 测试用例
                test_class = InputCellExcel(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()
            except Exception as e:
                print(e)

        if type_ == "按钮功能":
            self.pushButton_44.clicked.connect(
                lambda: os.startfile(self.comboBox_57.currentText())
            )
            self.pushButton_45.clicked.connect(
                lambda: self.merge_additional_functions("打开变量选择")
            )
            # 禁用中文输入
            self.lineEdit_28.setValidator(
                QRegExpValidator(QRegExp("[A-Za-z0-9]+"))
            )
            self.comboBox_57.activated.connect(
                lambda: self.find_controls("excel", "写入单元格")
            )
            # 测试按钮
            self.pushButton_59.clicked.connect(test)

        elif type_ == "写入参数":
            image, parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                image_=image,
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_57.clear()
            self.comboBox_57.addItems(extract_excel_from_global_parameter())
            self.find_controls("excel", "写入单元格")
        elif type_ == "还原参数":
            put_parameters(self.image_path, self.parameter_1)

    def ocr_recognition_function(self, type_):
        """ocr的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if (self.label_153.text() == "(0,0,0,0)") or (
                    self.comboBox_59.currentText() == ""
            ):
                QMessageBox.warning(self, "警告", "参数不能为空！")
                raise Exception
            parameter_dic_ = {
                "区域": self.label_153.text(),
                "变量": self.comboBox_59.currentText(),
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.label_153.setText(parameter_dic_.get("区域", "(0,0,0,0)"))
            self.comboBox_59.setCurrentText(parameter_dic_.get("变量", ""))

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
                parameter_dic_ = get_parameters()
                dic_ = self.get_test_dic(
                    repeat_number_=int(self.spinBox.value()),
                    parameter_1_=parameter_dic_,
                )

                # 测试用例
                client_info = get_ocr_info()
                if client_info["appId"] != "":
                    test_class = TextRecognition(self.out_mes, dic_)
                    test_class.is_test = True
                    test_class.start_execute()
                else:
                    QMessageBox.warning(self, "提示", "OCR未设置！")
                    open_setting_window()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f"指令错误请重试！", True)

        if type_ == "按钮功能":
            self.pushButton_48.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
            self.pushButton_46.clicked.connect(set_the_screenshot_area)
            self.pushButton_49.clicked.connect(open_setting_window)  # 打开百度ocr设置
            self.pushButton_47.clicked.connect(test)

        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_59.clear()
            self.comboBox_59.addItems(get_variable_info("list"))
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def get_mouse_position_function(self, type_):
        """获取鼠标位置的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            parameter_dic = {"变量": self.comboBox_60.currentText()}
            return parameter_dic

        def put_parameters(parameter_dic):
            """将参数还原到控件"""
            self.comboBox_60.setCurrentText(parameter_dic.get("变量", ""))

        if type_ == "按钮功能":
            self.pushButton_51.clicked.connect(
                lambda: self.merge_additional_functions("打开变量池")
            )
        elif type_ == "写入参数":
            parameter_1 = str(get_parameters())
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_1,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "加载信息":
            # 当t导航业显示时，加载信息到控件
            self.comboBox_60.clear()
            self.comboBox_60.addItems(get_variable_info("list"))
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def window_focus_wait_function(self, type_):
        """窗口焦点等待的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def get_parameters():
            """从tab页获取参数"""
            # 检查参数是否有异常
            if self.lineEdit_18.text() == "":
                QMessageBox.critical(self, "错误", "窗口标题未填！")
                raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                "标题包含": self.lineEdit_18.text(),
                "检测频率": self.comboBox_68.currentText(),
                "等待时间": self.spinBox_28.value(),
                "等待类型": self.comboBox_69.currentText()
            }
            return parameter_dic_

        def put_parameters(parameter_dic_):
            """将参数还原到控件"""
            self.lineEdit_18.setText(parameter_dic_["标题包含"])
            index = self.comboBox_68.findText(parameter_dic_["检测频率"])
            if index >= 0:
                self.comboBox_68.setCurrentIndex(index)
            self.spinBox_28.setValue(int(parameter_dic_["等待时间"]))
            index = self.comboBox_69.findText(parameter_dic_["等待类型"])
            if index >= 0:
                self.comboBox_69.setCurrentIndex(index)

        if type_ == "按钮功能":
            pass
        elif type_ == "加载信息":
            pass
        elif type_ == "写入参数":
            parameter_dic = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(
                instruction_=func_info_dic.get("指令类型"),
                repeat_number_=func_info_dic.get("重复次数"),
                exception_handling_=func_info_dic.get("异常处理"),
                parameter_1_=parameter_dic,
                remarks_=func_info_dic.get("备注"),
            )
        elif type_ == "还原参数":
            put_parameters(self.parameter_1)

    def color_judgment_function(self, type_):
        """颜色判断的功能
        :param self:
        :param type_: 功能名称（按钮功能、主要功能）"""

        def set_label_color():
            """设置标签颜色"""
            r = self.spinBox_26.value()
            g = self.spinBox_29.value()
            b = self.spinBox_30.value()
            color = f"background-color: rgb({r}, {g}, {b})"
            self.label_191.setStyleSheet(color)

        def open_color_picker():
            """打开颜色选择器"""
            color = QColorDialog.getColor()
            if color.isValid():
                self.spinBox_26.setValue(color.red())
                self.spinBox_29.setValue(color.green())
                self.spinBox_30.setValue(color.blue())
                set_label_color()

        def get_parameters(judge=False):
            """从tab页获取参数"""
            if not judge:
                if self.comboBox_74.currentText() == "" or self.comboBox_75.currentText() == "":
                    QMessageBox.critical(self, "错误", "分支未设置！")
                    raise ValueError
            # 返回参数字典
            parameter_dic_ = {
                '像素坐标': f"({self.label_197.text()}, {self.label_195.text()})",
                '目标颜色': f"({self.spinBox_26.value()}, {self.spinBox_29.value()}, {self.spinBox_30.value()})",
                '误差范围': self.spinBox_31.value(),
                '分支': f"{self.comboBox_74.currentText()}-{self.comboBox_75.currentText()}",
            }
            exception_handling_ = f"{self.comboBox_74.currentText()}-{self.comboBox_75.currentText()}"
            return parameter_dic_, exception_handling_

        def put_parameters(parameter_dic_):
            """将参数还原到tab页"""
            rgb_tuple = eval(parameter_dic_['目标颜色'])
            crosshair = eval(parameter_dic_['像素坐标'])
            self.label_197.setText(str(crosshair[0]))
            self.label_195.setText(str(crosshair[1]))
            self.spinBox_26.setValue(rgb_tuple[0])
            self.spinBox_29.setValue(rgb_tuple[1])
            self.spinBox_30.setValue(rgb_tuple[2])
            self.spinBox_31.setValue(int(parameter_dic_['误差范围']))
            self.comboBox_74.setCurrentText(parameter_dic_['分支'].split('-')[0])
            self.comboBox_75.setCurrentText(parameter_dic_['分支'].split('-')[1])

        def test():
            """测试功能"""
            try:
                parameter_dic_, exception_handling_ = get_parameters(True)
                dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                         parameter_1_=parameter_dic_
                                         )

                # 测试用例
                test_class = ColorJudgment(self.out_mes, dic_)
                test_class.is_test = True
                test_class.start_execute()

            except Exception as e:
                print(e)
                self.out_mes.out_mes(f'指令错误请重试！', True)

        if type_ == '按钮功能':
            self.pushButton_79.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "颜色判断"
                )
            )
            self.pushButton_82.pressed.connect(
                lambda: self.merge_additional_functions(
                    "change_get_mouse_position_function", "获取颜色"
                )
            )
            self.spinBox_26.valueChanged.connect(set_label_color)
            self.spinBox_29.valueChanged.connect(set_label_color)
            self.spinBox_30.valueChanged.connect(set_label_color)
            self.pushButton_80.clicked.connect(open_color_picker)
            # 分支选择
            self.comboBox_74.activated.connect(
                lambda: self.find_controls("分支", "颜色判断")
            )
            # 显示坐标
            self.toolButton_3.clicked.connect(
                lambda: pyautogui.moveTo(int(self.label_197.text()), int(self.label_195.text()))
            )
            # 测试按钮
            self.pushButton_81.clicked.connect(test)

        elif type_ == '写入参数':
            parameter_dic, exception_handling = get_parameters()
            # 将命令写入数据库
            func_info_dic = self.get_func_info()  # 获取功能区的参数
            self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                                 repeat_number_=func_info_dic.get('重复次数'),
                                                 exception_handling_=exception_handling,
                                                 parameter_1_=parameter_dic,
                                                 remarks_=func_info_dic.get('备注'))
        elif type_ == '加载信息':
            # 当t导航业显示时，加载信息到控件
            set_label_color()
            self.comboBox_74.clear()
            self.comboBox_74.addItems(get_branch_info(True))
            self.comboBox_74.setCurrentIndex(0)
            # 获取分支表名中的指令数量
            self.find_controls("分支", "颜色判断")

        elif type_ == '还原参数':
            put_parameters(self.parameter_1)
