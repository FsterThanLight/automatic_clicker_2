def convert_to_tuple(s):
    """将形如“(x, y)”的字符串转换为元组(x, y)。"""
    def convert_element(element):
        """
        尝试将字符串转换为适当的类型。
        - 如果可以转换为整数，返回整数。
        - 如果可以转换为浮点数，返回浮点数。
        - 否则，返回原字符串。
        """
        try:
            return int(element)
        except ValueError:
            try:
                return float(element)
            except ValueError:
                return element

    # 去除字符串中的圆括号和空格
    s = s.strip('()').replace(' ', '')
    # 以逗号分割字符串
    parts = s.split(',')
    if len(parts) == 2:
        # 将每个部分转换为合适的类型并返回元组
        return tuple(map(convert_element, parts))
    else:
        raise ValueError(f"输入格式不正确: {s}")


if __name__ == "__main__":
    # 测试示例
    try:
        tuple1 = convert_to_tuple("(0,0)")
        tuple2 = convert_to_tuple("(3.14, 2.718)")
        tuple3 = convert_to_tuple("(随机, 随机)")
        print(tuple1)  # 输出: (0.0, 0.0)
        print(tuple2)  # 输出: (3.14, 2.718)
        print(tuple3)  # 报错
    except ValueError as e:
        print(e)
