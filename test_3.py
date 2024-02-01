import pyautogui
import keyboard


def press_keyboard(key_):
    """鼠标移动事件
    :param key_: 按键列表"""
    # 将key全部转换为小写
    key = key_.lower()
    keys = key.split('+')
    # 按下键盘
    print('keys', keys)
    if len(keys) == 1:
        pyautogui.press(keys[0])  # 如果只有一个键,直接按下
    else:
        # 否则,组合多个键为热键
        hotkey = '+'.join(keys)
        print('hotkey', hotkey)
        pyautogui.hotkey(hotkey)
    # time.sleep(self.time_sleep)
    # self.main_window.plainTextEdit.appendPlainText('已经按下按键%s' % key)


if __name__ == '__main__':
    xxx = 'Ctrl+Alt+W'
    # xxx_l = xxx.lower()
    # print(xxx_l)
    # # hotkey = '+'.join(xxx)
    # pyautogui.hotkey(xxx_l)
    # 使用keyboard模块按下按键
    keyboard.press_and_release(xxx)
    keyboard.press_and_release('A')
