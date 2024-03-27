import pyautogui

# 获取鼠标当前位置
mouse_position = (pyautogui.position().x, pyautogui.position().y)
print(mouse_position)
print(type(mouse_position))
