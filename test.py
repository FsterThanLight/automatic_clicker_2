import re
import sys

list_all = [[(1, 2, 3, 4, ""), (4, 5, 6, 7, ""), (3, 4, 6, 8, "")],
            [(7, 8, 9, 2, ""), (3, 6, 9, 4, "3-1")],
            [(3, 2, 3, 4, ""), (3, 5, 6, 7, "4-1"), (3, 4, 6, 8, "")],
            [(4, 2, 3, 4, ""), (4, 5, 6, 7, ""), (4, 4, 6, 8, "")]]


def traverse_lists(current_list_index, current_index, list_all):
    while current_index < len(list_all[current_list_index]):
        elem = list_all[current_list_index][current_index]
        print(elem)
        print("执行函数")
        if elem[-1] == '':
            current_index += 1
        elif elem[-1] != '':
            branch_name_index, branch_index = elem[-1].split('-')
            x = int(branch_name_index) - 1
            y = int(branch_index) - 1
            traverse_lists(x, y, list_all)
            break


def get_a_number(cell_position, number=1):
    column_number = re.findall(r"[a-zA-Z]+", cell_position)[0]
    line_number = int(re.findall(r"\d+\.?\d*", cell_position)[0]) + number - 1
    new_cell_position = column_number + str(line_number)
    return new_cell_position


def string_judgment(filename):
    if filename.endswith('.py') or filename.endswith('.exe'):
        print('Python文件或可执行文件')
    else:
        print('未知文件类型')


if __name__ == '__main__':
    # cell_position = 'b2'
    # new_cell_position = get_a_number(cell_position, 4)
    # print(new_cell_position)
    # traverse_lists(0, 0, list_all)
    # x=input("输入任意字符结束")
    filename = r"C:\Users\federalsadler\Desktop\automatic_clicker_2\test.png"
    string_judgment(filename)
