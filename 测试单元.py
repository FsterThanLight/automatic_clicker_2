import unittest

from dateutil.parser import parse


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.assertEqual(
            self.comparison_variable('2024-02-23 23:45:28', '>', '2024-02-23 23:45:21'),
            False
        )

    @staticmethod
    def comparison_variable(variable1, comparison_symbol, variable2, variable_type):
        """比较变量"""

        def try_parse_date(variable):
            """尝试将变量解析为日期时间对象"""
            try:
                return parse(variable)
            except ValueError:
                return None

        variable1_ = variable1
        variable2_ = variable2
        if variable_type == '日期或时间':
            variable1_ = try_parse_date(variable1)
            variable2_ = try_parse_date(variable2)
        elif variable_type == '数字':
            variable1_ = eval(variable1)
            variable2_ = eval(variable2)
        elif variable_type == '字符串':
            variable1_ = str(variable1)
            variable2_ = str(variable2)

        if comparison_symbol == '=':
            return variable1_ == variable2_
        elif comparison_symbol == '≠':
            return variable1_ != variable2_
        elif comparison_symbol == '>':
            return variable1_ > variable2_
        elif comparison_symbol == '<':
            return variable1_ < variable2_
        elif comparison_symbol == '包含':
            return variable1_ in variable2_


if __name__ == '__main__':
    unittest.main()
