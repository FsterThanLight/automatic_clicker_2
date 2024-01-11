import time

import pyautogui


def wait_to_image(image, wait_instruction_type, timeout_period):
    """执行图片等待"""
    if wait_instruction_type == '等待到指定图像出现':
        # self.main_window.plainTextEdit.appendPlainText('正在等待指定图像出现中...')
        print('timeout_period', timeout_period)
        print('正在等待指定图像出现中...')
        # QApplication.processEvents()
        location = pyautogui.locateCenterOnScreen(
            image=image,
            confidence=0.8,
            minSearchTime=timeout_period
        )
        # print('location', location)
        # pyautogui.moveTo(location)
        if location:
            # self.main_window.plainTextEdit.appendPlainText('目标图像已经出现，等待结束')
            # QApplication.processEvents()
            print('目标图像已经出现，等待结束')

    elif wait_instruction_type == '等待到指定图像消失':
        vanish = True
        while vanish:
            try:
                location = pyautogui.locateCenterOnScreen(
                    image=image,
                    confidence=0.8,
                    minSearchTime=1
                )
                print('location', location)
            except pyautogui.ImageNotFoundException:
                print('目标图像已经消失，等待结束')
                vanish = False
            else:
                time.sleep(0.2)


if __name__ == '__main__':
    image = r'C:\Users\federalsadler\Desktop\xxx\test_4.png'
    image_2 = r'C:\Users\federalsadler\Desktop\xxx\test.png'
    wait_to_image(image, '等待到指定图像出现', 10)
    # wait_to_image(image_2, '等待到指定图像消失', 10)
