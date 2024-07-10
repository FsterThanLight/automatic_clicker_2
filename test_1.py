import win32clipboard


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


def get_clipboard_text():
    win32clipboard.OpenClipboard()
    try:
        text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    except Exception as e:
        print('获取剪贴板内容失败！', e)
        text = ''
    finally:
        win32clipboard.CloseClipboard()
    return text


if __name__ == "__main__":
    print(get_clipboard_text())
