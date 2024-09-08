import datetime
import json
import os
import time
import webbrowser
from contextlib import closing

import pymsgbox
import requests
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QDialog
from dateutil.parser import parse

from 软件信息 import INTERFACE, CURRENT_VERSION, DOWNLOAD_PAGE
from 窗体.update import Ui_Update_UI


class Check_Update(QThread):
    """自动更新"""
    show_update_signal = pyqtSignal(str, str, name='show_update_signal')
    show_update_window_signal = pyqtSignal(dict, name='show_update_window_signal')

    def __init__(self, parent=None):
        super(Check_Update, self).__init__(parent)
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
                "强制更新": data.get('forceUpdate'),
                "下载地址": data.get('download').split('，'),
                "完成后打开": data.get('完成后需要打开的文件', '').split('，'),
                "检测程序": data.get('检测程序', ''),
                "需要关闭的文件": data.get('需要关闭的文件', '').split('，'),
                "需要删除的文件": data.get('需要删除的文件', '').split('，'),
                "前往下载网页": data.get('前往下载网页', ''),
                "解压文件名": data.get('解压文件名', ''),
            }
        except Exception as e:
            print(e)
            return None

    def run(self):

        def open_update_window(update_info_dic_: dict):
            """打开更新窗口
            :param update_info_dic_: 更新信息字典
            """
            # 打开更新下载窗口
            self.show_update_window_signal.emit(update_info_dic_)

        update_info_dic = self.get_update_info()
        if update_info_dic is None:
            time.sleep(1)
            self.show_update_signal.emit('网络故障无法连接到服务器！\n\n请检查网络连接！', '错误')
            return

        new_version = update_info_dic.get('版本号')
        if CURRENT_VERSION == new_version:
            if self.is_show_info:
                self.show_update_signal.emit('当前已是最新版本！', '信息')
            return

        if update_info_dic.get('强制更新'):
            open_update_window(update_info_dic)
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
                if update_info_dic.get('前往下载网页') == '否':
                    open_update_window(update_info_dic)
                else:
                    webbrowser.open(DOWNLOAD_PAGE)


class Download_UpdatePack(QThread):
    """下载更新"""
    download_signal = pyqtSignal(str, name='download_signal')
    progress_signal = pyqtSignal(int, name='progress_signal')
    finish_signal = pyqtSignal(name='finish_signal')

    def __init__(self, parent=None):
        super(Download_UpdatePack, self).__init__(parent)
        self.download_url: str = ''

    def set_download_url(self, download_url_: str):
        """从外部设置下载地址"""
        self.download_url = download_url_

    def run(self):
        print('开始下载更新')
        self.download()
        # self.run_test()

    def download(self):
        """下载更新"""
        try:
            if (self.download_url is None) or (self.download_url == ''):
                return
            self.download_signal.emit('正在下载更新包...')
            with closing(requests.get(self.download_url, stream=True)) as response:
                chunk_size = 1024  # 单次请求最大值
                content_size = int(response.headers.get('content-length', 0))
                data_count = 0
                file_name = os.path.basename(self.download_url)  # 从链接中提取文件名
                # 下载文件，并显示self.progress进度条
                with open(file_name, "wb") as file:
                    for data in response.iter_content(chunk_size=chunk_size):
                        file.write(data)
                        data_count += len(data)
                        progress = int(data_count * 100 / content_size)
                        # 更新进度条
                        self.progress_signal.emit(progress)
            self.finish_signal.emit()
        except Exception as e:
            print(e)
            print('下载失败！请检查网络连接后重试！')
            self.download_signal.emit('下载失败！请检查网络连接后重试！')

    def run_test(self):
        try:
            if self.download_url is None:
                return
            self.download_signal.emit('正在下载更新包...')
            i = 0
            while i < 100:
                time.sleep(0.2)
                i += 1
                self.progress_signal.emit(i)
            self.finish_signal.emit()
        except Exception as e:
            print(e)
            print('下载失败！请检查网络连接后重试！')
            self.download_signal.emit('下载失败！请检查网络连接后重试！')


class UpdateWindow(QDialog, Ui_Update_UI):
    """更新窗口"""

    def __init__(self, parent=None, update_info_dic_: dict = None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)  # 隐藏帮助按钮
        self.update_info_dic = update_info_dic_
        self.progressBar.setValue(0)
        # 设置下载线程
        self.download_thread = Download_UpdatePack()
        self.download_thread.set_download_url(self.update_info_dic['下载地址'][0])
        self.download_thread.download_signal.connect(self.label_2.setText)
        self.download_thread.progress_signal.connect(self.progressBar.setValue)
        self.download_thread.finish_signal.connect(self.finish_download)
        # 开始下载
        self.download()

    def download(self):
        """下载更新"""
        # if self.download_thread.isRunning():
        self.download_thread.terminate()
        self.download_thread.start()

    def export_json(self):
        try:
            # 将 date 对象转换为字符串
            for key, value in self.update_info_dic.items():
                if isinstance(value, datetime.date):
                    self.update_info_dic[key] = value.isoformat()
            # 导出 JSON 文件
            with open('update_info.json', 'w', encoding='utf-8') as f:
                json.dump(self.update_info_dic, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"导出 JSON 文件时发生错误: {e}")

    def finish_download(self):
        """下载完成"""
        self.label_2.setText('下载完成！即将重启程序...')
        self.progressBar.setValue(100)
        self.export_json()  # 导出json文件
        time.sleep(1)
        try:
            os.startfile('sky.exe')
        except FileNotFoundError:
            self.label_2.setText('更新程序不存在！请手动解压更新文件！')
            os.remove('update_info.json')  # 删除json文件
        except OSError:
            self.label_2.setText(f'更新失败！请手动解压更新文件！')
            os.remove('update_info.json')

    def closeEvent(self, event):
        """关闭窗口时触发"""
        self.download_thread.terminate()
