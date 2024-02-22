import re
import time
import unittest


class MyTestCase(unittest.TestCase):
    def test_something(self):
        print(self.get_now_time())

    @staticmethod
    def get_now_time():
        """获取当前时间"""
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


if __name__ == '__main__':
    unittest.main()
