import datetime
import os
import sqlite3
import sys

from ini控制 import extract_resource_folder_path, get_branch_info

MAIN_FLOW = "主流程"


def timer(func):
    def func_wrapper(*args, **kwargs):
        from time import time

        time_start = time()
        result = func(*args, **kwargs)
        time_end = time()
        time_spend = time_end - time_start
        print("%s cost time: %.3f s" % (func.__name__, time_spend))
        return result

    return func_wrapper


def sqlitedb(db_name="命令集.db"):
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


def extract_excel_from_global_parameter():
    """从所有资源文件夹路径中提取所有的Excel文件
    :return: Excel文件列表"""
    # 从全局参数表中提取所有的资源文件夹路径
    resource_folder_path_list = extract_resource_folder_path()
    excel_files = []
    for folder_path in resource_folder_path_list:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if (
                            file.endswith(".xlsx") or file.endswith(".xls")
                    ) and not file.startswith("~$"):
                        excel_files.append(os.path.normpath(os.path.join(root, file)))
    return excel_files


def get_branch_count(branch_name: str) -> int:
    """获取分支表的数量
    :param branch_name: 分支表名
    :return: 目标分支表名中的指令数量"""
    # 连接数据库
    cursor, con = sqlitedb()
    # 获取表中数据记录的个数
    cursor.execute("SELECT count(*) FROM 命令 where 隶属分支=?", (branch_name,))
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
        cursor.execute("delete from 命令 where 隶属分支=?", (branch_name,))
    else:
        cursor.execute("delete from 命令 where ID<>-1")
    if judge:  # 清空全局参数表中所有的除了“主流程”的分支表名
        cursor.execute(
            "delete from 全局参数 " "where (分支表名 != ?  and 分支表名 is not null)",
            (MAIN_FLOW,),
        )
    con.commit()
    close_database(cursor, con)


def del_branch_in_database(branch_name):
    """删除数据库中的分支"""
    cursor, con = sqlitedb()
    cursor.execute(
        "delete from 命令 where 隶属分支=?", (branch_name,)
    )  # 从命令表中删除分支指令
    con.commit()
    close_database(cursor, con)  # 关闭数据库连接


def extracted_ins_from_database(branch_name=None) -> list:
    """从分支表中提取指令，如果不传入分支表名，则提取所有分支表中的指令
    :param branch_name: 分支表名，如果不传入，则提取所有指令
    :return: 分支表名列表"""

    def get_branch_table_ins(branch_name_: str) -> list:
        """获取某分支表名中的所有指令
        :param branch_name_ 目标分支表名
        :return 目标分支表名中的指令内容"""
        # 连接数据库
        cursor, con = sqlitedb()
        # 获取表中数据记录的个数
        cursor.execute("SELECT * FROM 命令 where 隶属分支=?", (branch_name_,))
        count_record = cursor.fetchall()
        # 关闭连接
        close_database(cursor, con)
        return count_record

    # 提取所有分支中的指令
    if branch_name:
        return get_branch_table_ins(branch_name)  # 返回分支指令列表
    else:
        # 提取所有分支表中的指令
        branch_table_name_list = get_branch_info(keys_only=True)
        all_list_instructions = []
        if len(branch_table_name_list) != 0:
            for branch_table_name in branch_table_name_list:
                all_list_instructions.append(get_branch_table_ins(branch_table_name))
            return all_list_instructions


def extracted_ins_target_id_from_database(id_: int) -> list:
    """获取目标id的指令，并返回一个和extracted_ins_from_database相似的列表
    :param id_: 目标id"""
    cursor, con = sqlitedb()
    cursor.execute("SELECT * FROM 命令 where ID=?", (id_,))
    count_record = cursor.fetchall()
    close_database(cursor, con)
    # 生成一个和extracted_ins_from_database相似的列表
    return [count_record]


# @timer
def writes_to_recently_opened_files(file_path: str):
    """将最近打开的文件写入数据库
    :param file_path: 文件路径"""

    def write_to_new_file(cursor_, file_path_, time_stamp_) -> None:
        # 查找数据库中是否存在该文件路径,如果存在则更新打开时间，如果不存在则插入数据
        cursor_.execute("SELECT * FROM 最近打开 WHERE 文件路径 = ?", (file_path_,))
        result = cursor_.fetchone()
        if result:
            cursor_.execute(
                "UPDATE 最近打开 SET 打开时间=? WHERE 文件路径 = ?",
                (time_stamp_, file_path_),
            )
        else:
            cursor_.execute(
                "INSERT INTO 最近打开(文件路径, 打开时间) VALUES (?, ?)",
                (file_path_, time_stamp_),
            )

    def delete_the_oldest_file(cursor_, con_, keep_number=10) -> None:
        """从数据库中删除最早的文件"""
        try:
            cursor_.execute("SELECT 文件路径 FROM 最近打开 ORDER BY 打开时间 ")
            result_ = cursor_.fetchall()

            if len(result_) > keep_number:
                # 只保留最近打开的3个文件
                files_to_keep = [item[0] for item in result_[-keep_number:]]
                print(files_to_keep)
                # 根据文件路径删除记录
                cursor_.execute(
                    "DELETE FROM 最近打开"
                    " WHERE 文件路径 not IN ({})".format(
                        ",".join("?" * len(files_to_keep))
                    ),
                    files_to_keep,
                )
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


def get_recently_opened_file(judge="单文件"):
    """获取最近打开的文件
    :param judge: 返回类型（单文件、文件列表）
    :return: 最近打开的文件"""
    cursor, con = sqlitedb()
    cursor.execute("SELECT 文件路径 FROM 最近打开 ORDER BY 打开时间 DESC")
    result = cursor.fetchall()
    close_database(cursor, con)
    if judge == "单文件":
        return os.path.normpath([item[0] for item in result][0])
    elif judge == "文件列表":
        return [item[0] for item in result]


def remove_recently_opened_file(file_path: str):
    """从最近打开的文件中删除指定的文件
    :param file_path: 文件路径"""
    cursor, con = sqlitedb()
    cursor.execute("DELETE FROM 最近打开 WHERE 文件路径 = ?", (file_path,))
    con.commit()
    close_database(cursor, con)


def get_value_from_variable_table():
    """从设置表中获取指定设置类型的值
    :return: 设置类型的值"""
    cursor, con = sqlitedb()
    cursor.execute("SELECT * FROM 变量池")
    result = cursor.fetchall()
    close_database(cursor, con)
    return result


def set_value_to_variable_table(variable_list: list):
    """将指定设置类型的值写入变量池窗口的表格
    :param variable_list: 将要写入的变量列表（变量名称、备注、变量值）"""
    cursor, con = sqlitedb()
    # 查询数据库中的现有值
    try:
        cursor.execute("SELECT * FROM 变量池")
        existing_values = cursor.fetchall()
        # 将现有值存储为字典，便于比较
        existing_values_dict = {row[0]: (row[1], row[2]) for row in existing_values}
        # 遍历传入的变量列表
        for variable_name, remark, value in variable_list:
            # 如果变量名称在数据库中已存在且对应的备注值不等于传入值，则更新备注值
            if variable_name in existing_values_dict:
                cursor.execute(
                    "UPDATE 变量池 SET 备注 = ?, 值 = ? WHERE 变量名称 = ?",
                    (remark, value, variable_name),
                )
            # 如果变量名称不在数据库中，则插入新的记录
            elif variable_name not in existing_values_dict:
                cursor.execute(
                    "INSERT INTO 变量池(变量名称, 备注, 值) VALUES (?, ?, ?)",
                    (variable_name, remark, value),
                )
                cursor.execute(
                    "UPDATE 变量池 SET 值 = ? WHERE 变量名称 = ?",
                    (value, variable_name),
                )
        # 检查变量池中是否有未在传入变量列表中的变量，如果有，则删除这些记录
        for variable_name in existing_values_dict:
            if variable_name not in [v[0] for v in variable_list]:
                cursor.execute(
                    "DELETE FROM 变量池 WHERE 变量名称 = ?", (variable_name,)
                )
        con.commit()
    except sqlite3.IntegrityError:
        print("An error occurred: 数据库中已存在该变量名称")
    finally:
        close_database(cursor, con)


def get_variable_info(return_type: str):
    """从变量名中获取变量信息，可以选择返回类型为字典或列表
    :param return_type: 指定返回类型，'dict'表示返回字典，'list'表示返回列表"""
    cursor, conn = sqlitedb()
    try:
        if return_type == "dict":
            cursor.execute(f"SELECT 变量名称, 值 FROM 变量池")
            result = {
                item[0]: item[1] for item in cursor.fetchall()
            }  # 获取变量名称和值的字典
        elif return_type == "list":
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
        cursor.execute(
            "UPDATE 变量池 SET 值 = ? WHERE 变量名称 = ?", (new_value, variable_name)
        )
        conn.commit()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        close_database(cursor, conn)


if __name__ == "__main__":
    print(extracted_ins_from_database())
    print(extracted_ins_target_id_from_database(13))
