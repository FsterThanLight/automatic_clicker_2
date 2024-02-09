import os
import sqlite3
import sys

MAIN_FLOW = '主流程'


def timer(func):
    def func_wrapper(*args, **kwargs):
        from time import time
        time_start = time()
        result = func(*args, **kwargs)
        time_end = time()
        time_spend = time_end - time_start
        print('%s cost time: %.3f s' % (func.__name__, time_spend))
        return result

    return func_wrapper


def sqlitedb(db_name='命令集.db'):
    """建立与数据库的连接，返回游标
    :param db_name: 数据库名称
    :return: 游标，数据库连接"""
    try:
        # 取得当前文件目录
        con = sqlite3.connect(db_name)
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


def get_setting_data_from_db(*args):
    """从数据库中获取设置参数
    :param args: 设置类型参数
    :return: 设置参数字典"""
    cursor, conn = sqlitedb()
    if len(args) > 1:
        placeholders = ','.join(['?' for _ in args])
        query = f'SELECT 设置类型, 值 FROM 设置 WHERE 设置类型 IN ({placeholders})'
        cursor.execute(query, args)
        results = cursor.fetchall()
        close_database(cursor, conn)
        settings_dict = {setting_type: value for setting_type, value in results}
        return settings_dict
    elif len(args) == 1:
        cursor.execute('SELECT 值 FROM 设置 WHERE 设置类型 = ?', (args[0],))
        result = cursor.fetchone()
        close_database(cursor, conn)  # 关闭数据库
        return result[0] if result else None
    else:
        return None


# @timer
def update_settings_in_database(**kwargs):
    """在数据库中更新指定表中的设置类型的值
    :param kwargs: 设置类型和对应值的关键字参数，如：暂停时间=1, 时间间隔=1, 图像匹配精度=0.8
    """
    if kwargs:
        try:
            cursor, conn = sqlitedb()
            for setting_type, value in kwargs.items():
                query = f"UPDATE 设置 SET 值=? WHERE 设置类型 = ?"
                cursor.execute(query, (value, setting_type))
            conn.commit()
            close_database(cursor, conn)
        except sqlite3.Error as e:
            print(f"Error updating database: {e}")


# 全局参数的数据库操作
def global_write_to_database(judge, value):
    """将全局参数写入数据库
    :param judge: 判断写入类型（资源文件夹路径、分支表名）
    :param value: 资源文件夹路径"""
    # 连接数据库
    cursor, conn = sqlitedb()
    if judge == '资源文件夹路径':
        cursor.execute('INSERT INTO 全局参数(资源文件夹路径,分支表名) VALUES (?,?)',
                       (value, None))
        conn.commit()
    elif judge == '分支表名':
        if value != MAIN_FLOW:
            cursor, con = sqlitedb()
            cursor.execute(
                'insert into 全局参数(资源文件夹路径,分支表名) '
                'values(?,?)',
                (None, value)
            )
            con.commit()
    close_database(cursor, conn)


def extract_global_parameter(column_name: str) -> list:
    """从全局参数表中提取指定列的数据
    :param column_name: 列名（资源文件夹路径、分支表名）"""
    cursor, conn = sqlitedb()
    cursor.execute(f"select {column_name} from 全局参数")
    # 去除None并转换为列表
    result_list = [item[0] for item in cursor.fetchall() if item[0] is not None]
    close_database(cursor, conn)
    return result_list


def extract_excel_from_global_parameter():
    """从所有资源文件夹路径中提取所有的Excel文件
    :return: Excel文件列表"""
    # 从全局参数表中提取所有的资源文件夹路径
    resource_folder_path_list = extract_global_parameter('资源文件夹路径')
    excel_files = []
    for folder_path in resource_folder_path_list:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.endswith('.xlsx') or file.endswith('.xls'):
                        excel_files.append(os.path.normpath(os.path.join(root, file)))
    return excel_files


def get_branch_count(branch_name: str) -> int:
    """获取分支表的数量
    :param branch_name: 分支表名
    :return: 目标分支表名中的指令数量"""
    # 连接数据库
    cursor, con = sqlitedb()
    # 获取表中数据记录的个数
    cursor.execute('SELECT count(*) FROM 命令 where 隶属分支=?', (branch_name,))
    count_record = cursor.fetchone()[0]
    # 关闭连接
    close_database(cursor, con)
    return count_record


def clear_all_ins(judge: bool = False, branch_name: str = None):
    """清空数据库中所有指令
    :param judge: 是否清除分支表名
    :param branch_name: 分支表名，如果不传入，则清空所有分支表名的数据"""
    cursor, con = sqlitedb()
    # 清空分支列表中所有的数据
    if branch_name:  # 清空指定分支表名的数据
        cursor.execute('delete from 命令 where 隶属分支=?', (branch_name,))
    else:
        cursor.execute('delete from 命令 where ID<>-1')
    if judge:  # 清空全局参数表中所有的除了“主流程”的分支表名
        cursor.execute(
            'delete from 全局参数 '
            'where (分支表名 != ?  and 分支表名 is not null)',
            (MAIN_FLOW,)
        )
    con.commit()
    close_database(cursor, con)


def save_window_size(save_size: tuple, window_name: str):
    """获取窗口大小
    :param save_size: 保存时的窗口大小
    :param window_name:（主窗口、设置窗口、导航窗口）
    :return: 窗口大小"""
    cursor, con = sqlitedb()
    # 查找数据库中是否有该设置类型
    cursor.execute('SELECT * FROM 设置 WHERE 设置类型 = ?', (window_name,))
    result = cursor.fetchone()
    if result:
        cursor.execute('UPDATE 设置 SET 值=? WHERE 设置类型 = ?', (str(save_size), window_name))
    else:
        cursor.execute('INSERT INTO 设置(设置类型, 值) VALUES (?, ?)', (window_name, str(save_size)))
    con.commit()
    close_database(cursor, con)


def set_window_size(window_name: str):
    """设置窗口大小
    :param window_name:（主窗口、设置窗口、导航窗口）
    :return: 窗口大小"""
    try:
        height, width = eval(get_setting_data_from_db(window_name))
        return int(height), int(width)
    except TypeError:
        return 0, 0


def extracted_ins_from_database(branch_name=None) -> list:
    """提取所有分支表名
    :param branch_name: 分支表名，如果不传入，则提取所有指令
    :return: 分支表名列表"""

    def get_branch_table_ins(branch_name_: str) -> list:
        """获取某分支表名中的所有指令
        :param branch_name_ 目标分支表名
        :return 目标分支表名中的指令内容"""
        # 连接数据库
        cursor, con = sqlitedb()
        # 获取表中数据记录的个数
        cursor.execute('SELECT * FROM 命令 where 隶属分支=?', (branch_name_,))
        count_record = cursor.fetchall()
        # 关闭连接
        close_database(cursor, con)
        return count_record

    # 提取所有分支中的指令
    if branch_name:
        return get_branch_table_ins(branch_name)  # 返回分支指令列表
    else:
        # 提取所有分支表中的指令
        branch_table_name_list = extract_global_parameter('分支表名')
        all_list_instructions = []
        if len(branch_table_name_list) != 0:
            for branch_table_name in branch_table_name_list:
                all_list_instructions.append(get_branch_table_ins(branch_table_name))
            return all_list_instructions


if __name__ == '__main__':
    pass
    all_list_instructions_ = extracted_ins_from_database()
    print(all_list_instructions_)
    print(len(all_list_instructions_))
    for i in all_list_instructions_:
        print(i)
    # print(get_branch_table_ins('分支1'))
