import keyboard

xxx = 'w'


def print_pressed_keys(e):
    global xxx
    if e.name == xxx:
        print(f'{e.name} was pressed')
    else:
        print(f'other key was pressed')


if __name__ == '__main__':
    keyboard.on_press(print_pressed_keys)
    # 持续监听按键，直到按下 'esc' 键退出
    keyboard.wait('esc')
