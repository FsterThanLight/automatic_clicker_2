import base64
import json

import requests


class YdmVerify(object):
    _custom_url = "http://api.jfbym.com/api/YmServer/customApi"
    _token = "C_5EONXWamEz12ED1Z4oZyxDQCvwr1hg--ww-SUMumI"
    _headers = {
        'Content-Type': 'application/json'
    }

    def common_verify(self, image, verify_type="60000"):
        # 数英汉字类型
        # 通用数英1-4位 10110
        # 通用数英5-8位 10111
        # 通用数英9~11位 10112
        # 通用数英12位及以上 10113
        # 通用数英1~6位plus 10103
        # 定制-数英5位~qcs 9001
        # 定制-纯数字4位 193
        # 中文类型
        # 通用中文字符1~2位 10114
        # 通用中文字符 3~5位 10115
        # 通用中文字符6~8位 10116
        # 通用中文字符9位及以上 10117
        # 定制-XX西游苦行中文字符 10107
        # 计算类型
        # 通用数字计算题 50100
        # 通用中文计算题 50101
        payload = {
            "image": base64.b64encode(image).decode(),
            "token": self._token,
            "type": verify_type
        }
        # print(payload)
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        # print(resp.text)
        return resp.json()['data']['data']


if __name__ == '__main__':
    y = YdmVerify()
    image_path = r'C:\Users\FS\Downloads\test_2.png'
    image = open(image_path, 'rb').read()
    # 使用base64编码
    print(y.common_verify(image, verify_type="10110"))

