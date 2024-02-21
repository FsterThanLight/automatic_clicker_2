import re
import unittest


class MyTestCase(unittest.TestCase):
    def test_something(self):
        pass
        # self.assertEqual(True, False)
        or_text = 'jilejghl;kjal;iejooookf;sldkg@@hjieg@@k;oe@@xxx@@jli'

        value_dict_ = {  # 修改此处
            'hjieg': '经济',
            'xxx': '德国'
        }

        xxx = self.substitution_variable(or_text, value_dict_)
        print(xxx)

    @staticmethod
    def substitution_variable(text: str, value_dict: dict) -> str:
        """将text中的@@variable@@替换为value"""
        for key, value in value_dict.items():
            text = text.replace(f'@@{key}@@', value)
        return text


if __name__ == '__main__':
    unittest.main()
