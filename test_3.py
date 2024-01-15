import os
import keyboard as keyboard
from time import sleep
from PIL import Image, ImageGrab

if __name__ == '__main__':
    while True:
        if keyboard.wait(hotkey='ctrl + alt + a') is None:
            # 清空剪切板
            from ctypes import windll
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()

            print('开始截图')
            os.system('start /B rundll32 PrScrn.dll PrScrn')

            # 等待截图后放到剪切板
            im = ImageGrab.grabclipboard()
            while not im:
                im = ImageGrab.grabclipboard()
                sleep(0.5)

            print('截图完成')
            if isinstance(im, Image.Image):
                im.save('Picture.png')
                print('保存成功')
            else:
                print('保存失败，重新截图')

