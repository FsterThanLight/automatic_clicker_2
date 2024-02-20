import time

import win32con
import win32gui


def show_normal_window_with_specified_title(title, judge='最大化'):
    """将指定标题的窗口置顶
    :param title: 指定标题
    :param judge: 判断（最大化、最小化、显示窗口、关闭）"""

    def get_all_window_title():
        """获取所有窗口句柄和窗口标题"""
        hwnd_title_ = dict()

        def get_all_hwnd(hwnd, mouse):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                hwnd_title_.update({hwnd: win32gui.GetWindowText(hwnd)})

        win32gui.EnumWindows(get_all_hwnd, 0)
        return hwnd_title_

    hwnd_title = get_all_window_title()
    for h, t in hwnd_title.items():
        if title in t:
            if judge == '最大化':
                win32gui.ShowWindow(h, win32con.SW_SHOWMAXIMIZED)  # 最大化显示窗口
            elif judge == '最小化':
                win32gui.ShowWindow(h, win32con.SW_SHOWMINIMIZED)
            elif judge == '显示窗口':
                win32gui.ShowWindow(h, win32con.SW_SHOWNORMAL)  # 显示窗口
            elif judge == '关闭窗口':
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            break
    else:
        print(f'没有找到标题为“{title}”的窗口！')


if __name__ == '__main__':
    time.sleep(5)
    show_normal_window_with_specified_title('Clash', '关闭')
