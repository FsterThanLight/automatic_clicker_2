import os

from 数据库操作 import extract_global_parameter


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


@timer
def get_available_path(image_name_: str):
    """组合图片路径，返回可以打开的图片路径
    :param image_name_: 图片路径或者图片名称
    :return: 可以打开的图片路径"""

    def search_image_in_folders(image_name_only_, folders):
        for folder_path in folders:
            image_path = os.path.join(folder_path, image_name_only_)
            if os.path.exists(image_path):
                # print('找到图片路径：', image_path)
                return image_path
        return None

    if os.path.isabs(image_name_):
        if os.path.exists(image_name_):
            # print('图片路径已经存在')
            return image_name_
        else:
            # print('传入的图片路径不存在，尝试重新匹配路径')
            image_name_only = os.path.basename(image_name_)
            res_folder_path = extract_global_parameter('资源文件夹路径')
            return search_image_in_folders(image_name_only, res_folder_path)

    else:
        res_folder_path = extract_global_parameter('资源文件夹路径')
        return search_image_in_folders(image_name_, res_folder_path)


if __name__ == '__main__':
    # image_name = '9v3IsrOkm1.png'
    image_name = r'C:\Users\FS\Desktop\Clicker-test\9v3IsrOkm1.png'
    # 组合图片路径，返回可以打开的图片路径
    print(get_available_path(image_name))
