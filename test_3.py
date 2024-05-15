import ctypes
import random


def random_position(self, position, random_range):
    """设置随机坐标"""
    if random_range == 0:
        return position
    x, y = position
    x_random = random.randint(-random_range, random_range)
    y_random = random.randint(-random_range, random_range)
    return x + x_random, y + y_random


if __name__ == "__main__":
    new_pos = random_position(None, (100, 500), 0)
    print(new_pos)
