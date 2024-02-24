xxx = 1
python_code = ("import time\n"
               "'xxx'.replace(',','')\n"  # 执行代码
               "")  # 返回result变量的值

if __name__ == '__main__':
    # 1. 使用exec()执行python代码
    try:
        # 定义全局命名空间字典
        globals_dict = {}
        # 在执行代码时，将结果保存到全局命名空间中
        exec(python_code, globals_dict)
        # 从全局命名空间中获取结果
        result = globals_dict.get('result', None)
        print("Result:", result)  # 输出结果
    except Exception as e:
        print(e)
