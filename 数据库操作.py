import datetime
import os
import re
import sqlite3
import sys
import time

import win32con
import win32gui
import winsound

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


def get_str_now_time():
    """获取当前时间"""
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def line_number_increment(old_value, number=1):
    """行号递增
    :param old_value: 旧的单元格号
    :param number: 递增的数量"""
    # 提取字母部分和数字部分
    column_letters = re.findall(r"[a-zA-Z]+", old_value)[0]
    line_number = int(re.findall(r"\d+\.?\d*", old_value)[0])
    # 计算新的行号
    new_line_number = line_number + number
    # 组合字母部分和新的行号
    new_cell_position = (column_letters + str(new_line_number)).upper()
    new_cell_position = new_cell_position
    return new_cell_position


def show_normal_window_with_specified_title(title):
    """将指定标题的窗口正常显示"""

    def get_window_titles(hwnd, titles):
        titles[hwnd] = win32gui.GetWindowText(hwnd)

    if eval(get_setting_data_from_db('任务完成后显示主窗口')):
        hwnd_title = {}
        win32gui.EnumWindows(get_window_titles, hwnd_title)

        for h, t in hwnd_title.items():
            if t == title:
                try:
                    time.sleep(0.5)
                    # # 设置窗口样式为可见、不透明
                    # win32gui.SetWindowLong(
                    #     h, win32con.GWL_EXSTYLE,
                    #     win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE) & ~win32con.WS_EX_LAYERED
                    # )
                    win32gui.ShowWindow(h, win32con.SW_SHOWNORMAL)  # 正常显示窗口
                    # win32gui.SetForegroundWindow(h)
                except Exception as e:
                    print(f"主窗口显示出现错误: {e}")
                break


def system_prompt_tone(judge: str):
    """系统提示音
    :param judge: 判断类型（线程结束、全局快捷键、执行异常）"""
    is_tone = eval(get_setting_data_from_db('系统提示音'))
    if judge == '线程结束' and is_tone:
        for i_ in range(3):
            winsound.Beep(500, 300)
    elif judge == '全局快捷键' and is_tone:
        winsound.Beep(500, 300)
    elif judge == '执行异常' and is_tone:
        winsound.Beep(1000, 1000)


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


# @timer
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


def set_window_size(window):
    def get_window_size(window_name: str):
        """设置窗口大小
        :param window_name:（主窗口、设置窗口、导航窗口）
        :return: 窗口大小"""
        try:
            height_, width_ = eval(get_setting_data_from_db(window_name))
            return int(height_), int(width_)
        except TypeError:
            return 0, 0

    width, height = get_window_size(window.windowTitle())
    if width and height:
        window.resize(width, height)


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


# @timer
def writes_to_recently_opened_files(file_path: str):
    """将最近打开的文件写入数据库
    :param file_path: 文件路径"""

    def write_to_new_file(cursor_, file_path_, time_stamp_) -> None:
        # 查找数据库中是否存在该文件路径,如果存在则更新打开时间，如果不存在则插入数据
        cursor_.execute('SELECT * FROM 最近打开 WHERE 文件路径 = ?',
                        (file_path_,))
        result = cursor_.fetchone()
        if result:
            cursor_.execute('UPDATE 最近打开 SET 打开时间=? WHERE 文件路径 = ?',
                            (time_stamp_, file_path_))
        else:
            cursor_.execute('INSERT INTO 最近打开(文件路径, 打开时间) VALUES (?, ?)',
                            (file_path_, time_stamp_))

    def delete_the_oldest_file(cursor_, con_, keep_number=10) -> None:
        """从数据库中删除最早的文件"""
        try:
            cursor_.execute('SELECT 文件路径 FROM 最近打开 ORDER BY 打开时间 ')
            result_ = cursor_.fetchall()

            if len(result_) > keep_number:
                # 只保留最近打开的3个文件
                files_to_keep = [item[0] for item in result_[-keep_number:]]
                print(files_to_keep)
                # 根据文件路径删除记录
                cursor_.execute(
                    'DELETE FROM 最近打开'
                    ' WHERE 文件路径 not IN ({})'.format(','.join('?' * len(files_to_keep))),
                    files_to_keep)
            else:
                print("数据库中没有足够的文件需要删除")
        except Exception as e_:
            print("An error occurred:", e_)
        finally:
            con_.commit()

    # 将时间转化为13位时间戳
    time_stamp = int(datetime.datetime.now().timestamp() * 1000)
    try:
        # 连接数据库
        cursor, con = sqlitedb()
        write_to_new_file(cursor, file_path, time_stamp)
        delete_the_oldest_file(cursor, con)  # 删除最早打开的文件
        con.commit()
        close_database(cursor, con)
    except Exception as e:
        print("An error occurred:", e)


def get_recently_opened_file(judge='单文件'):
    """获取最近打开的文件
    :param judge: 返回类型（单文件、文件列表）
    :return: 最近打开的文件"""
    cursor, con = sqlitedb()
    cursor.execute('SELECT 文件路径 FROM 最近打开 ORDER BY 打开时间 DESC')
    result = cursor.fetchall()
    close_database(cursor, con)
    if judge == '单文件':
        return os.path.normpath([item[0] for item in result][0])
    elif judge == '文件列表':
        return [item[0] for item in result]


def remove_recently_opened_file(file_path: str):
    """从最近打开的文件中删除指定的文件
    :param file_path: 文件路径"""
    cursor, con = sqlitedb()
    cursor.execute('DELETE FROM 最近打开 WHERE 文件路径 = ?', (file_path,))
    con.commit()
    close_database(cursor, con)


def get_value_from_variable_table():
    """从设置表中获取指定设置类型的值
    :return: 设置类型的值"""
    cursor, con = sqlitedb()
    cursor.execute('SELECT * FROM 变量池')
    result = cursor.fetchall()
    close_database(cursor, con)
    return result


def set_value_to_variable_table(variable_list: list):
    """将指定设置类型的值写入变量池窗口的表格
    :param variable_list: 将要写入的变量列表（变量名称、备注、变量值）"""
    if len(variable_list) != 0:
        cursor, con = sqlitedb()
        # 查询数据库中的现有值
        try:
            cursor.execute('SELECT * FROM 变量池')
            existing_values = cursor.fetchall()
            # 将现有值存储为字典，便于比较
            existing_values_dict = {row[0]: (row[1], row[2]) for row in existing_values}
            # 遍历传入的变量列表
            for variable_name, remark, value in variable_list:
                # 如果变量名称在数据库中已存在且对应的备注值不等于传入值，则更新备注值
                if variable_name in existing_values_dict:
                    cursor.execute(
                        'UPDATE 变量池 SET 备注 = ?, 值 = ? WHERE 变量名称 = ?',
                        (remark, value, variable_name))
                # 如果变量名称不在数据库中，则插入新的记录
                elif variable_name not in existing_values_dict:
                    cursor.execute(
                        'INSERT INTO 变量池(变量名称, 备注, 值) VALUES (?, ?, ?)',
                        (variable_name, remark, value))
                    cursor.execute(
                        'UPDATE 变量池 SET 值 = ? WHERE 变量名称 = ?',
                        (value, variable_name))
            # 检查变量池中是否有未在传入变量列表中的变量，如果有，则删除这些记录
            for variable_name in existing_values_dict:
                if variable_name not in [v[0] for v in variable_list]:
                    cursor.execute(
                        'DELETE FROM 变量池 WHERE 变量名称 = ?',
                        (variable_name,))
            con.commit()
        except sqlite3.IntegrityError:
            print("An error occurred: 数据库中已存在该变量名称")
        finally:
            close_database(cursor, con)


def get_variable_info(return_type: str) -> dict or list:
    """从变量名中获取变量信息，可以选择返回类型为字典或列表
    :param return_type: 指定返回类型，'dict'表示返回字典，'list'表示返回列表"""
    cursor, conn = sqlitedb()
    try:
        if return_type == 'dict':
            cursor.execute(f"SELECT 变量名称, 值 FROM 变量池")
            result = {item[0]: item[1] for item in cursor.fetchall()}  # 获取变量名称和值的字典
        elif return_type == 'list':
            cursor.execute(f"SELECT 变量名称 FROM 变量池")
            result = [item[0] for item in cursor.fetchall()]  # 获取变量名称的列表
        else:
            raise ValueError("Invalid return_type. Use 'dict' or 'list'.")
    except Exception as e:
        print(f"An error occurred: {e}")
        result = None
    finally:
        close_database(cursor, conn)
    return result


def set_variable_value(variable_name, new_value) -> None:
    """设置变量池中的变量的值
    :param variable_name: 变量名称
    :param new_value: 新的值"""
    cursor, conn = sqlitedb()
    try:
        cursor.execute("UPDATE 变量池 SET 值 = ? WHERE 变量名称 = ?", (new_value, variable_name))
        conn.commit()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        close_database(cursor, conn)


if __name__ == '__main__':
    print(get_variable_info('list'))
    print(get_variable_info('dict'))
    # set_variable_value('xx', 'ji')
