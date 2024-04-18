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


if __name__ == '__main__':
    pass
    # image_name = '9v3IsrOkm1.png'
    # image_name = r'C:\Users\FS\Desktop\Clicker-test\9v3IsrOkm1.png'
    # # 组合图片路径，返回可以打开的图片路径
    # print(get_available_path(image_name))
