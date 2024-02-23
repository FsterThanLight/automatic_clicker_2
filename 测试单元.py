import re
import unittest

import openpyxl


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.assertEqual(self.line_number_increment('a1', 3), 'A4')

    def line_number_increment(self, old_value, number=1):
        """行号递增
        :param old_value: 旧的单元格号
        :param number: 递增的数量"""
        # 提取字母部分和数字部分
        column_letters = re.findall(r"[a-zA-Z]+", old_value)[0]
        line_number = int(re.findall(r"\d+\.?\d*", old_value)[0])
        # 计算新的行号
        new_line_number = line_number + number
        # 组合字母部分和新的行号
        new_cell_position = (column_letters + str(new_line_number)).upper()
        new_cell_position = new_cell_position
        return new_cell_position


if __name__ == '__main__':
    unittest.main()
