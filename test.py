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
            x = int(branch_name_index)-1
            y = int(branch_index)-1
            traverse_lists(x, y, list_all)
            break


if __name__ == '__main__':
    traverse_lists(0, 0, list_all)
    x=input("输入任意字符结束")
