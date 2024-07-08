import datetime
import re
import time
import typing

import winsound
import win32con
import win32gui
from system_hotkey import SystemHotkey, user32

from ini控制 import get_setting_data_from_ini


def get_str_now_time():
    """获取当前时间"""
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def line_number_increment(old_value, number=1):
    """行号递增
    :param old_value: 旧的单元格号
    :param number: 递增的数量"""
    # 提取字母部分和数字部分
    column_letters = re.findall(r"[a-zA-Z]+", old_value)[0]
    line_number = int(re.findall(r"\d+\.?\d*", old_value)[0])
    # 计算新的行号
    new_line_number = line_number + number
    # 组合字母部分和新的行号
    new_cell_position = (column_letters + str(new_line_number)).upper()
    new_cell_position = new_cell_position
    return new_cell_position


def show_normal_window_with_specified_title(title):
    """将指定标题的窗口正常显示"""

    def get_window_titles(hwnd, titles):
        titles[hwnd] = win32gui.GetWindowText(hwnd)

    if eval(get_setting_data_from_ini('Config', '任务完成后显示主窗口')):
        hwnd_title = {}
        win32gui.EnumWindows(get_window_titles, hwnd_title)

        for h, t in hwnd_title.items():
            if t == title:
                try:
                    time.sleep(0.5)
                    win32gui.ShowWindow(h, win32con.SW_SHOWNORMAL)  # 正常显示窗口
                except Exception as e:
                    print(f"主窗口显示出现错误: {e}")
                break


def system_prompt_tone(judge: str):
    """系统提示音
    :param judge: 判断类型（线程结束、全局快捷键、执行异常）"""
    try:
        is_tone = eval(get_setting_data_from_ini('Config', '系统提示音'))
        if judge == '线程结束' and is_tone:
            for i_ in range(3):
                winsound.Beep(500, 300)
        elif judge == '全局快捷键' and is_tone:
            winsound.Beep(500, 300)
        elif judge == '执行异常' and is_tone:
            winsound.Beep(1000, 1000)
    except Exception as e:
        print('系统提示音错误！', e)


def is_hotkey_valid(hkobj: SystemHotkey, hk: typing.List[str]):
    """判断快捷键是否有效"""
    hk = hkobj.order_hotkey(hk)
    try:
        keycode, masks = hkobj.parse_hotkeylist(hk)
        reg_hk_res = user32.RegisterHotKey(None, 1, masks, keycode)
        if reg_hk_res:
            user32.UnregisterHotKey(None, reg_hk_res)
            return True
    except Exception as e:
        print("获取快捷键注册信息失败！", e)
    return False


def show_window(title):
    """将指定标题的窗口正常显示，主要用于主窗口显示"""
    def get_window_titles(hwnd, titles):
        titles[hwnd] = win32gui.GetWindowText(hwnd)

    try:
        hwnd_title = {}
        win32gui.EnumWindows(get_window_titles, hwnd_title)
        for h, t in hwnd_title.items():
            if t == title:
                win32gui.ShowWindow(h, win32con.SW_SHOWNORMAL)  # 正常显示窗口
                win32gui.SetForegroundWindow(h)
                break
    except Exception as e:
        print(f"显示窗口出现错误: {e}")
