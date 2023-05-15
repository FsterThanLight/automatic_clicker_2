import sys

list_1 = [(1, 2, 3, 4, ""), (4, 5, 6, 7, "list_2-1"), (3, 4, 6, 8, "")]
list_2 = [(7, 8, 9, 2, ""), (3, 6, 9, 4, "list_3-0")]
list_3 = [(3, 2, 3, 4, ""), (3, 5, 6, 7, "list_4-0"), (3, 4, 6, 8, "")]
list_4 = [(4, 2, 3, 4, ""), (4, 5, 6, 7, ""), (4, 4, 6, 8, "")]
all_lists = [list_1, list_2, list_3, list_4]


def traverse_lists(current_list, current_index, *args):
    while current_index < len(current_list):
        elem = current_list[current_index]
        print(elem)
        print("执行函数")
        if elem[-1] == '':
            current_index += 1
        elif elem[-1] != '':
            branch_name, branch_index = elem[-1].split('-')
            current_list = eval(branch_name)
            current_index = int(branch_index)
            traverse_lists(current_list, current_index, *args)
            break


if __name__ == '__main__':
    traverse_lists(list_1, 0, *all_lists)
