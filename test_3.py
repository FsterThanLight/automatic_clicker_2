from 数据库操作 import get_variable_info


def sub_variable(text: str, is_str: bool = False):
    """将text中的变量替换为变量值"""
    new_text = text
    if ('☾' in text) and ('☽' in text):
        variable_dic = get_variable_info('dict')
        for key, value in variable_dic.items():
            if is_str:
                new_text = new_text.replace(f'☾{key}☽', str(value))
            else:
                new_text = new_text.replace(f'☾{key}☽', str(f'"{value}"'))
    return new_text


def sub_variable_2(text: str):
    """将text中的变量替换为变量值"""
    new_text = text
    if ('☾' in text) and ('☽' in text):
        variable_dic = get_variable_info('dict')
        for key, value in variable_dic.items():
            new_text = new_text.replace(f'☾{key}☽', str(f'"{value}"'))
    return new_text


if __name__ == '__main__':
    # 1. 使用exec()执行python代码
    python_code = sub_variable_2("☾俄国☽.replace('x','y')")
    print(python_code)
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
