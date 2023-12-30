import sqlite3
import sys


def sqlitedb():
    """建立与数据库的连接，返回游标
    :return: 游标，数据库连接"""
    try:
        # 取得当前文件目录
        con = sqlite3.connect('命令集.db')
        cursor = con.cursor()
        return cursor, con
    except sqlite3.Error:
        print("未连接到数据库！！请检查数据库路径是否异常。")
        sys.exit()


def close_database(cursor, conn):
    """关闭数据库
    :param cursor: 游标
    :param conn: 数据库连接"""
    cursor.close()
    conn.close()


def get_setting_data_from_db() -> tuple:
    """从数据库中获取设置参数
    :return: 设置参数"""
    cursor, conn = sqlitedb()
    cursor.execute('select * from 设置')
    list_setting_data = cursor.fetchall()
    close_database(cursor, conn)
    # 使用字典来存储设置参数
    setting_dict = {i[0]: i[1] for i in list_setting_data}
    return (setting_dict.get('持续时间'),
            setting_dict.get('时间间隔'),
            setting_dict.get('图像匹配精度'),
            setting_dict.get('暂停时间'))
