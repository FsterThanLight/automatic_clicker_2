import os
import sys

import pypyodbc


class SettingsData:
    def __init__(self):
        self.duration = 0
        self.interval = 0
        self.confidence = 0
        self.time_sleep = 0

    def init(self, oddbc_name):
        """设置初始化"""
        # 从数据库加载设置
        """建立与数据库的连接，返回游标"""
        # 取得当前文件目录
        cursor, conn = self.accdb(oddbc_name)
        # 从数据库中取出全部数据
        cursor.execute('select * from 设置')
        # 读取全部数据
        list_setting_data = cursor.fetchall()
        # 关闭连接
        self.close_database(cursor, conn)
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

    def accdb(self, name):
        """建立与数据库的连接，返回游标"""
        try:
            path = os.path.abspath('.')
            # 取得当前文件目录
            mdb = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path + '\\' + name
            print(mdb)
            # 连接字符串
            conn = pypyodbc.win_connect_mdb(mdb)
            # 建立连接
            cursor = conn.cursor()
            return cursor, conn
        except pypyodbc.Error:
            x = input("未连接到数据库！！请检查数据库路径是否异常。")
            sys.exit()

    def close_database(self, cursor, conn):
        """关闭数据库"""
        cursor.close()
        conn.close()


if __name__ == '__main__':
    odbc_name = '命令集.accdb'
    setting = SettingsData()
    setting.init(odbc_name)
    print(setting.confidence)
