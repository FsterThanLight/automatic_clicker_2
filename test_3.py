import random

import pyautogui


def mouse_moves(direction, distance):
    """鼠标移动事件"""
    # 相对于当前位置移动鼠标
    directions = {'↑': (0, -1), '↓': (0, 1), '←': (-1, 0), '→': (1, 0)}
    if direction in directions:
        x, y = directions.get(direction)
        pyautogui.moveRel(x * int(distance), y * int(distance), duration=0.25)


def mouse_moves_random_1():
    """鼠标移动事件"""
    # 相对于当前位置移动鼠标
    directions = {'↑': (0, -1), '↓': (0, 1), '←': (-1, 0), '→': (1, 0)}
    direction = random.choice(list(directions.keys()))
    if direction in directions:
        x, y = directions.get(direction)
        distance = random.randint(1, 500)
        duration_ran = random.uniform(0.1, 0.9)
        try:
            pyautogui.moveRel(x * distance, y * distance, duration=duration_ran)
        except pyautogui.FailSafeException:
            pass


def mouse_moves_random_2():
    """鼠标移动事件"""
    screen_width, screen_height = pyautogui.size()
    # 随机生成坐标
    x = random.randint(0, screen_width)
    y = random.randint(0, screen_height)
    # 随机生成时间
    duration_ran = random.uniform(0.1, 0.9)
    try:
        pyautogui.moveTo(x, y, duration=duration_ran)
    except pyautogui.FailSafeException:
        pass


if __name__ == '__main__':
    i = 0
    while i < 10:
        # mouse_moves_random_2()
        mouse_moves_random()
        i += 1
