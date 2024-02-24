import unittest


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.assertEqual(True, False)
        # 获取当前时间
        # now = datetime.datetime.now()


if __name__ == '__main__':
    unittest.main()
