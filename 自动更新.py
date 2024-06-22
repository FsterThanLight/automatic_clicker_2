import datetime
import os
import time

import requests
from dateutil.parser import parse

INTERFACE = 'http://api.ytsoftware.cn/appClicker/getClickerAppVersion'


class Update:
    def __init__(self, parent=None):
        self.parent = parent

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
        os.startfile('update.exe')


# if __name__ == '__main__':
#     print(Update.get_update_info())
    # Update.start_update_program()
