import pyautogui


def move_mouse_to_coordinates(x: int, y: int):
    print(f"Moving mouse to coordinates {x}, {y}")
    # 使用pyautogui库实现鼠标移动
    pyautogui.moveTo(x, y, duration=10)


if __name__ == "__main__":
    move_mouse_to_coordinates(4000, 100)
