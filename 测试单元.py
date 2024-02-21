import unittest

import keyboard


class MyTestCase(unittest.TestCase):
    def test_something(self):
        # self.assertEqual(True, False)
        self.key_waits()

    def key_waits(self):
        """按键等待"""
        print('按键等待')
        keyboard.wait('w')


if __name__ == '__main__':
    unittest.main()
