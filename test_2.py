import os


def run_external_file(file_path):
    """运行外部文件"""
    try:
        os.startfile(file_path)
        print(f'运行成功：{file_path}')
    except Exception as e:
        print(e)


if __name__ == '__main__':
    file_1 = r'D:\待看影视\黑色孤儿 第四季\黑色孤儿第四季08.mp'
    run_external_file(file_1)
