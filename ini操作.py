import configparser


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


def get_setting_data_from_ini(selection: str, *args):
    """从ini文件中获取设置数据
    :param selection: 选择的选区域
    :param args: 设置类型参数"""
    try:
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        if len(args) > 1:
            setting_data_dic = {}
            for arg in args:
                setting_data_dic[arg] = config[selection][arg]
            return setting_data_dic
        elif len(args) == 1:
            return config[selection][args[0]]
        else:
            return None
    except Exception as e:
        print(e)
        return None


def update_settings_in_ini(selection: str, **kwargs):
    """更新ini文件中的设置数据
    :param selection: 选择的选区域
    :param kwargs: 设置类型参数"""
    try:
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        for key, value in kwargs.items():
            config[selection][key] = value
        with open("config.ini", "w", encoding="utf-8") as f:
            config.write(f)
    except Exception as e:
        print("更新设置数据失败！", e)


def get_ocr_info() -> dict:
    """从ini中获取百度OCR的API信息"""
    try:
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        return {
            "appId": config["三方接口"]["appid"],
            "apiKey": config["三方接口"]["apikey"],
            "secretKey": config["三方接口"]["secretkey"]
        }
    except Exception as e:
        print("获取OCR信息失败！", e)
        return {}


def save_window_size(save_size: tuple, window_name: str):
    """获取窗口大小
    :param save_size: 保存时的窗口大小
    :param window_name:（主窗口、设置窗口、导航窗口）
    :return: 窗口大小"""
    try:
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        # 检查'窗口大小'选区中是否存在window_name选项
        config["窗口大小"][window_name] = str(save_size)
        with open("config.ini", "w", encoding="utf-8") as f:
            config.write(f)
    except Exception as e:
        print("保存窗口大小失败！", e)


def set_window_size(window):
    def get_window_size(window_name: str):
        """设置窗口大小
        :param window_name:（主窗口、设置窗口、导航窗口）
        :return: 窗口大小"""
        try:
            height_, width_ = eval(get_setting_data_from_ini("窗口大小", window_name))
            return int(height_), int(width_)
        except TypeError:
            return 0, 0

    width, height = get_window_size(window.windowTitle())
    if width and height:
        window.resize(width, height)


if __name__ == "__main__":
    print(get_ocr_info())
