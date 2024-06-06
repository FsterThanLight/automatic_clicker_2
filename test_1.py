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


import configparser


@timer
def test():
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    print(config['Config']['图像匹配精度'])


if __name__ == "__main__":
    # 读取配置文件
    test()
