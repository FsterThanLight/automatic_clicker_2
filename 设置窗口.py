from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QMessageBox, QDialog
from system_hotkey import SystemHotkey

from functions import is_hotkey_valid
from ini操作 import (
    update_settings_in_ini,
    get_setting_data_from_ini,
    set_window_size,
    save_window_size, get_global_shortcut, set_global_shortcut)
from 窗体.setting import Ui_Setting

BAIDU_OCR = 'https://ai.baidu.com/tech/ocr'
YUN_CODE = 'https://www.jfbym.com'


class Setting(QDialog, Ui_Setting):
    """添加设置窗口"""

    def __init__(self, parent=None):
        super().__init__(parent)
        # 初始化设置窗口
        self.setupUi(self)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)  # 隐藏帮助按钮
        set_window_size(self)  # 获取上次退出时的窗口大小
        # 绑定事件
        self.parent().unregister_global_shortcut_keys()  # 注销全局快捷键
        self.pushButton.clicked.connect(self.save_setting)  # 点击保存（应用）按钮
        self.pushButton_3.clicked.connect(self.restore_default)  # 点击恢复至默认按钮
        self.pushButton_2.clicked.connect(lambda: self.open_link(BAIDU_OCR))  # 打开百度OCR链接
        self.pushButton_4.clicked.connect(lambda: self.open_link(YUN_CODE))  # 打开云码链接
        self.radioButton_2.clicked.connect(lambda: self.change_mode('极速模式'))  # 开启极速模式
        self.radioButton.clicked.connect(lambda: self.change_mode('普通模式'))  # 切换普通模式
        self.load_setting_data()  # 加载设置数据

    def save_setting_date(self):
        """保存设置数据"""
        # 模式选择
        model = self.radioButton.text() if self.radioButton.isChecked() else \
            self.radioButton_2.text() if self.radioButton_2.isChecked() else None
        # 更新ini文件
        update_settings_in_ini(
            'Config',
            图像匹配精度=str(self.horizontalSlider.value() / 10),
            时间间隔=str(self.horizontalSlider_2.value() / 1000),
            持续时间=str(self.horizontalSlider_3.value() / 1000),
            暂停时间=str(self.horizontalSlider_4.value() / 1000),
            模式=model,
            启动检查更新=str(True if self.checkBox.isChecked() else False),
            退出提醒清空指令=str(True if self.checkBox_2.isChecked() else False),
            系统提示音=str(True if self.checkBox_3.isChecked() else False),
            任务完成后显示主窗口=str(True if self.checkBox_4.isChecked() else False)
        )
        update_settings_in_ini(
            '三方接口',
            appId=str(self.lineEdit.text()),
            apiKey=str(self.lineEdit_2.text()),
            secretKey=str(self.lineEdit_3.text()),
            云码Token=str(self.lineEdit_6.text())
        )

        # 更新快捷键设置，检查快捷键是否有效，无效则弹出提示
        def validate_and_set_hotkey(hotkey, key_sequence_edit_, action_):
            """验证并设置快捷键"""
            key_sequence = key_sequence_edit_.keySequence().toString().lower().split('+')
            key_sequence = [key.replace('ctrl', 'control') for key in key_sequence]
            if is_hotkey_valid(hotkey, key_sequence):
                set_global_shortcut(**{action_: key_sequence})
            else:
                QMessageBox.information(
                    self, '提醒',
                    f'快捷键{key_sequence_edit_.keySequence().toString()}为无效按键！'
                    f'\n\n可能的原因：'
                    f'\n1.系统不支持注册的按键。'
                    f'\n2.按键已被其他程序占用。'
                )
                raise Exception('无效的快捷键！')

        key_mapping = {
            '开始运行': self.keySequenceEdit,
            '结束运行': self.keySequenceEdit_2,
            '分支选择': self.keySequenceEdit_3,
            '暂停和恢复': self.keySequenceEdit_4
        }
        for action, key_sequence_edit in key_mapping.items():
            validate_and_set_hotkey(SystemHotkey(), key_sequence_edit, action)

    def save_setting(self):
        """保存按钮事件"""
        try:
            self.save_setting_date()
            QMessageBox.information(self, '提醒', '设置已经生效！')
            # 退出设置窗口
            self.close()
        except Exception as e:
            print('保存设置失败！', e)

    def restore_default(self):
        """设置恢复至默认"""
        self.radioButton.isChecked()
        self.horizontalSlider.setValue(8)
        self.horizontalSlider_2.setValue(200)
        self.horizontalSlider_3.setValue(200)
        self.horizontalSlider_4.setValue(100)

    def load_setting_data(self):
        """加载设置数据库中的数据"""
        # 加载设置数据
        setting_data_dic = get_setting_data_from_ini(
            'Config',
            '图像匹配精度',
            '时间间隔',
            '持续时间',
            '暂停时间',
            '模式',
            '启动检查更新',
            '退出提醒清空指令',
            '系统提示音',
            '任务完成后显示主窗口'
        )
        app_data_dic = get_setting_data_from_ini(
            '三方接口',
            'appId',
            'apiKey',
            'secretKey',
            '云码Token'
        )
        self.horizontalSlider.setValue(int(float(setting_data_dic['图像匹配精度']) * 10))
        self.horizontalSlider_2.setValue(int(float(setting_data_dic['时间间隔']) * 1000))
        self.horizontalSlider_3.setValue(int(float(setting_data_dic['持续时间']) * 1000))
        self.horizontalSlider_4.setValue(int(float(setting_data_dic['暂停时间']) * 1000))

        if setting_data_dic['模式'] == '极速模式':
            self.radioButton_2.setChecked(True)
            self.change_mode('极速模式')
        else:
            self.radioButton.setChecked(True)
            self.change_mode('普通模式')

        self.checkBox.setChecked(eval(setting_data_dic['启动检查更新']))
        self.checkBox_2.setChecked(eval(setting_data_dic['退出提醒清空指令']))
        self.checkBox_3.setChecked(eval(setting_data_dic['系统提示音']))
        self.checkBox_4.setChecked(eval(setting_data_dic['任务完成后显示主窗口']))

        # 填入OCR API信息
        self.lineEdit.setText(app_data_dic['appId'])
        self.lineEdit_2.setText(app_data_dic['apiKey'])
        self.lineEdit_3.setText(app_data_dic['secretKey'])

        # 填入云码Token
        self.lineEdit_6.setText(app_data_dic['云码Token'])

        # 加载快捷键设置
        global_shortcut_dic = get_global_shortcut()
        self.keySequenceEdit.setKeySequence('+'.join(global_shortcut_dic['开始运行']))
        self.keySequenceEdit_2.setKeySequence('+'.join(global_shortcut_dic['结束运行']))
        self.keySequenceEdit_3.setKeySequence('+'.join(global_shortcut_dic['分支选择']))
        self.keySequenceEdit_4.setKeySequence('+'.join(global_shortcut_dic['暂停和恢复']))

    def change_mode(self, mode: str):
        """切换模式
        :param mode: 模式（极速模式、普通模式）"""
        if mode == '极速模式':
            self.horizontalSlider_2.setValue(0)
            self.horizontalSlider_3.setValue(100)
            self.horizontalSlider_4.setValue(0)
            self.horizontalSlider_2.setEnabled(False)
            self.horizontalSlider_4.setEnabled(False)
            self.pushButton_3.setEnabled(False)
        elif mode == '普通模式':
            self.horizontalSlider_2.setEnabled(True)
            self.horizontalSlider_4.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.restore_default()

    @staticmethod
    def open_link(url):
        """打开网页"""
        QDesktopServices.openUrl(QUrl(url))

    def closeEvent(self, event):
        # 窗口大小
        save_window_size((self.width(), self.height()), self.windowTitle())
        # 重新注册全局快捷键
        self.parent().register_global_shortcut_keys()
