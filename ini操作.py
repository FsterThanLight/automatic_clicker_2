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


def get_config() -> configparser.ConfigParser:
    """获取配置文件"""
    config = configparser.ConfigParser()
    config.read("config.ini", encoding="utf-8")
    return config


def get_setting_data_from_ini(selection: str, *args):
    """从ini文件中获取设置数据
    :param selection: 选择的选区域
    :param args: 设置类型参数"""
    try:
        config = get_config()
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
        config = get_config()
        for key, value in kwargs.items():
            config[selection][key] = value
        with open("config.ini", "w", encoding="utf-8") as f:
            config.write(f)
    except Exception as e:
        print("更新设置数据失败！", e)


def get_ocr_info() -> dict:
    """从ini中获取百度OCR的API信息"""
    try:
        config = get_config()
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
        config = get_config()
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


def get_global_shortcut():
    """获取全局快捷键"""
    try:
        config = get_config()
        return {
            "开始运行": config["全局快捷键"]["开始运行"].lower().split("+"),
            "结束运行": config["全局快捷键"]["结束运行"].lower().split("+"),
            "分支选择": config["全局快捷键"]["分支选择"].lower().split("+"),
            "暂停和恢复": config["全局快捷键"]["暂停和恢复"].lower().split("+"),
        }
    except Exception as e:
        print("获取全局快捷键失败！", e)
        return {}


def set_global_shortcut(**kwargs):
    """设置全局快捷键
    :param kwargs: 全局快捷键参数, 如：开始运行=["ctrl", "alt", "1"]"""
    try:
        config = get_config()
        for key, value in kwargs.items():
            # 将"control"替换为"ctrl"
            value = ['ctrl' if v.lower() == 'control' else v for v in value]
            config["全局快捷键"][key] = "+".join(value).lower()
        with open("config.ini", "w", encoding="utf-8") as f:
            config.write(f)
    except Exception as e:
        print("设置全局快捷键失败！", e)


def writes_to_resource_folder_path(path: str):
    """将资源文件路径写入到config.ini中"""
    try:
        config = get_config()
        section = '资源文件夹路径'
        if not config.has_section(section):
            config.add_section(section)
        paths = {key: config.get(section, key) for key in config.options(section)}
        if path not in paths.values():
            keys = sorted(
                [int(
                    k.replace("路径", "")
                ) for k in paths.keys() if k.replace("路径", "").isdigit()], reverse=True
            )
            new_key = f"路径{keys[0] + 1 if keys else 1}"
            config.set(section, new_key, path)
            with open("config.ini", "w", encoding="utf-8") as configfile:
                config.write(configfile)
            print("路径写入成功！")
        else:
            print("路径已经存在！")
    except Exception as e:
        print("写入资源文件路径失败！", e)


def del_resource_folder_path(path: str):
    """删除资源文件路径"""
    try:
        config = get_config()
        section = '资源文件夹路径'
        # 检查 '资源文件夹路径' 部分是否存在
        if not config.has_section(section):
            print("配置文件中不存在资源文件夹路径部分！")
            return
        # 获取所有路径键值对并检查路径是否存在
        paths = {key: config.get(section, key) for key in config.options(section)}
        if path not in paths.values():
            print("路径不存在于配置文件中！")
            return
        # 删除指定路径并重新编号
        new_paths = {f"路径{i + 1}": value for i, value in enumerate(v for k, v in paths.items() if v != path)}
        # 清除原有部分并重新添加整理后的路径键值对
        config.remove_section(section)
        config.add_section(section)
        for key, value in new_paths.items():
            config.set(section, key, value)
        # 保存配置文件
        with open("config.ini", "w", encoding="utf-8") as configfile:
            config.write(configfile)
        print("路径删除成功！")
    except Exception as e:
        print("删除资源文件路径失败！", e)


def extract_resource_folder_path() -> list:
    """提取资源文件夹路径"""
    try:
        config = get_config()
        section = '资源文件夹路径'
        if not config.has_section(section):
            return []
        paths = {key: config.get(section, key) for key in config.options(section)}
        return list(paths.values())
    except Exception as e:
        print("提取资源文件夹路径失败！", e)
        return []


if __name__ == "__main__":
    # writes_to_resource_folder_path("C:/Users/zhuzh/Desktop/11")
    # writes_to_resource_folder_path("C:/Users/zhuzh/Desktop/12")
    # writes_to_resource_folder_path("C:/Users/zhuzh/Desktop/13")
    # writes_to_resource_folder_path("C:/Users/zhuzh/Desktop/14")
    # del_resource_folder_path("C:/Users/zhuzh/Desktop/12")
    print(extract_resource_folder_path())
