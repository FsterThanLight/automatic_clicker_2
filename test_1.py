import time
import pygetwindow as gw


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


def check_focus(window_title_: str, timeout: int = 10, frequency: float = 0.5, wait_for_focus: bool = True):
    """检查窗口是否获得焦点
    :param window_title_: 窗口标题
    :param timeout: 超时时间
    :param frequency: 检查频率
    :param wait_for_focus: True表示等待窗口获取焦点，False表示等待窗口失去焦点"""
    start_time = time.time()
    while True:
        active_window = gw.getActiveWindow()
        if active_window is not None:
            if wait_for_focus:
                if window_title_ in active_window.title:
                    print("应用程序已经获得了焦点")
                    break
            else:
                if window_title_ not in active_window.title:
                    print("应用程序已经失去焦点")
                    break
        else:
            print("没有找到活动窗口")

        # 检查超时
        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            raise TimeoutError("超过指定时间未获取到焦点" if wait_for_focus else "超过指定时间未失去焦点")

        time.sleep(frequency)


if __name__ == "__main__":
    check_focus('Typora', 5)
