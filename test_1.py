import psutil
import pyautogui
import uiautomation as auto


def get_pid(name):
    """  
    作用：根据进程名获取进程pid
    返回：返回匹配第一个进程的pid
    """
    pids = psutil.process_iter()
    for pid in pids:
        if pid.name() == name:
            return pid.pid


if __name__ == "__main__":
    print(get_pid('WeChat.exe'))
    pyautogui.hotkey('ctrl', 'alt', 'w')  # 打开微信窗口
    wei_xin = auto.WindowControl(searchDepth=1, ClassName='WeChatMainWndForPC')
    print(wei_xin)
    wx_chat_win = wei_xin.ListItemControl(searchDepth=10, Name='文件传输助手')
    wx_chat_win.Click()
