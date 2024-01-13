import pyautogui
import pyperclip
import win32gui
import win32process
from pywinauto import Application


# @timer
def check_course(title_):
    """检查蒙泰软件是否正在运行
    :param title_: 窗口标题"""

    def get_all_window_title():
        """获取所有窗口句柄和窗口标题"""
        hwnd_title_ = dict()

        def get_all_hwnd(hwnd_, mouse):
            # print(mouse)
            if win32gui.IsWindow(hwnd_) and win32gui.IsWindowEnabled(hwnd_) and win32gui.IsWindowVisible(hwnd_):
                hwnd_title_.update({hwnd_: win32gui.GetWindowText(hwnd_)})

        win32gui.EnumWindows(get_all_hwnd, 0)
        return hwnd_title_

    hwnd_title = get_all_window_title()
    for h, t in hwnd_title.items():
        if t == title_:
            return h


def send_message_to_wechat(contact_person, message):
    """向微信好友发送消息
    :param contact_person: 联系人
    :param message: 消息内容"""

    def get_process_id(hwnd_):
        thread_id, process_id_ = win32process.GetWindowThreadProcessId(hwnd_)
        return process_id_

    hwnd = check_course('微信')
    if hwnd:
        pyautogui.hotkey('ctrl', 'alt', 'w')  # 打开微信窗口
        process_id = get_process_id(hwnd)  # 获取微信进程id
        # 连接到wx
        wx_app = Application(backend='uia').connect(process=process_id)
        # 定位到主窗口
        wx_win = wx_app.window(class_name='WeChatMainWndForPC')
        wx_chat_win = wx_win.child_window(title=contact_person, control_type="ListItem")
        # 聚焦到所需的对话框
        wx_chat_win.click_input()
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        # 模拟按下键盘enter键
        pyautogui.press('enter')
    else:
        print('未找到微信窗口')


if __name__ == '__main__':
    # send_message_to_wechat('文件传输助手', '测试')
    # 从剪切板获取数据

    # 获取剪切板上的内容
    clipboard_content = pyperclip.paste()

    # 输出内容
    print("Clipboard Content:", clipboard_content)

