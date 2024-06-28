import datetime
import os
import time

import pymsgbox
import requests
from PyQt5.QtCore import QThread, pyqtSignal
from dateutil.parser import parse

CURRENT_VERSION = "v0.26.2 Beta"
INTERFACE = 'http://api.ytsoftware.cn/appClicker/getClickerAppVersion'


class Update(QThread):
    """自动更新"""
    show_update_signal = pyqtSignal(str, str, name='show_update_signal')

    def __init__(self, parent=None):
        super(Update, self).__init__(parent)
        self.parent = parent
        self.is_show_info = False

    def set_show_info(self, is_show_info: bool):
        self.is_show_info = is_show_info

    @staticmethod
    def get_update_info():
        """获取更新信息"""

        def time_13_to_date(timestamp) -> datetime:
            """获取13位时间戳中的日期"""
            if len(str(timestamp)) == 13:  # 如果是13位时间戳,则转换成日期
                timeArray = time.localtime(int(timestamp) / 1000)
                otherStyleDate = datetime.datetime(*timeArray[:6])
                return otherStyleDate.date()
            elif len(str(timestamp)) == 19:  # 如果是正常的时间格式,则转换成日期
                normal_date = parse(timestamp).date()
                return normal_date

        try:
            headers = {'Content-Type': 'application/json'}
            res = requests.post(INTERFACE, headers=headers, timeout=5)
            data = res.json().get('data')
            return {
                "版本号": data.get('version'),
                "更新内容": data.get('desc'),
                "更新时间": time_13_to_date(data.get('updateTime')),
                "强制更新": data.get('forceUpdate')
            }
        except Exception as e:
            print(e)
            return None

    @staticmethod
    def start_update_program():
        """启动更新程序"""
        try:
            os.startfile('update.exe')
        except FileNotFoundError:
            pymsgbox.alert(
                text='更新程序不存在！请重新下载',
                title='更新提示',
                icon=pymsgbox.STOP
            )

    def run(self):
        update_info_dic = self.get_update_info()

        if update_info_dic is None:
            pymsgbox.alert(
                text='网络故障无法连接到服务器！\n\n请检查网络连接！',
                title='更新提示',
                icon=pymsgbox.STOP
            )
            self.show_update_signal.emit('网络故障无法连接到服务器！\n\n请检查网络连接！', '错误')
            return

        new_version = update_info_dic.get('版本号')
        print(f'当前版本：{CURRENT_VERSION}，最新版本：{new_version}')
        if CURRENT_VERSION == new_version:
            if self.is_show_info:
                self.show_update_signal.emit('当前已是最新版本！', '信息')
            return

        if update_info_dic.get('强制更新'):
            self.start_update_program()
        else:
            text = (f"发现新版本：{new_version}，"
                    f"\n\n{update_info_dic['更新内容']}"
                    f"\n\n是否更新？")
            reply = pymsgbox.confirm(
                text=text,
                title='更新提示',
                icon=pymsgbox.INFO,
                buttons=[pymsgbox.OK_TEXT, pymsgbox.CANCEL_TEXT]
            )
            if reply == pymsgbox.OK_TEXT:
                self.start_update_program()
