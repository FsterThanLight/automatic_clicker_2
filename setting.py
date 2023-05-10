import sqlite3


class SettingsData:
    def __init__(self):
        self.duration = 0
        self.interval = 0
        self.confidence = 0
        self.time_sleep = 0

    def init(self):
        """设置初始化"""
        # 从数据库加载设置
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        cursor.execute('select * from 设置')
        list_setting_data = cursor.fetchall()
        con.close()
        # print(list_setting_data)

        for i in range(len(list_setting_data)):
            if list_setting_data[i][0] == '图像匹配精度':
                self.confidence = list_setting_data[i][1]
            elif list_setting_data[i][0] == '时间间隔':
                self.interval = list_setting_data[i][1]
            elif list_setting_data[i][0] == '持续时间':
                self.duration = list_setting_data[i][1]
            elif list_setting_data[i][0] == '暂停时间':
                self.time_sleep = list_setting_data[i][1]


if __name__ == '__main__':
    setting = SettingsData()
    setting.init()
