import tkinter as tk
import pyautogui


class ScreenCapture:
    def __init__(self):
        self.rect = None  # 截图的矩形区域
        # 鼠标左键按下的位置
        self.x_1 = 0
        self.y_1 = 0
        # 鼠标左键抬起的位置
        self.x_3 = 0
        self.y_3 = 0
        self.pic = None

    def screenshot_area(self):
        """截取屏幕矩形区域"""
        root = tk.Tk()
        root.attributes('-fullscreen', True)
        root.attributes('-alpha', 0.3)
        root.configure(bg='grey')

        # 鼠标左键按下事件
        def on_press(event):
            self.x_1, self.y_1 = event.x, event.y
            print('鼠标开始点击位置为：', self.x_1, self.y_1)
            self.rect = canvas.create_rectangle(self.x_1, self.y_1, 0, 0, outline='red', width=2, fill='black')

        def on_drag(event):
            x_2, y_2 = event.x, event.y
            canvas.coords(self.rect, self.x_1, self.y_1, x_2, y_2)
            # 将矩形区域内的颜色设置为透明
            root.wm_attributes("-transparentcolor", "black")

        def on_release(event):
            self.x_3, self.y_3 = event.x, event.y
            print('鼠标释放位置为：', self.x_3, self.y_3)
            canvas.delete(self.rect)
            canvas.destroy()
            root.destroy()

        # 创建屏幕遮罩
        canvas = tk.Canvas(root, bg="grey", cursor='cross')
        canvas.pack(fill=tk.BOTH, expand=True)
        # 绑定鼠标事件
        canvas.bind('<ButtonPress-1>', on_press)
        canvas.bind('<B1-Motion>', on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        # 开始事件循环
        root.mainloop()
        self.pic = pyautogui.screenshot(region=(self.x_1, self.y_1, self.x_3 - self.x_1, self.y_3 - self.y_1))


if __name__ == '__main__':
    screen_capture = ScreenCapture()
    screen_capture.screenshot_area()
