import base64
import json
import time

import requests


class YdmVerify(object):
    _custom_url = "http://api.jfbym.com/api/YmServer/customApi"
    _token = ""
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
        # 定制-计算题 cni 452
        payload = {
            "image": base64.b64encode(image).decode(),
            "token": self._token,
            "type": verify_type
        }
        print(payload)
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def slide_verify(self, slide_image, background_image, verify_type="20101"):
        # 滑块类型
        # 通用双图滑块  20111
        payload = {
            "slide_image": base64.b64encode(slide_image).decode(),
            "background_image": base64.b64encode(background_image).decode(),
            "token": self._token,
            "type": verify_type
        }

        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def sin_slide_verify(self, image, verify_type="20110"):
        # 通用单图滑块(截图)  20110
        payload = {
            "image": base64.b64encode(image).decode(),
            "token": self._token,
            "type": verify_type
        }
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def traffic_slide_verify(self, seed, data, href, verify_type="900010"):
        # 定制-滑块协议slide_traffic  900010
        payload = {
            "seed": seed,
            "data": data,
            "href": href,
            "token": self._token,
            "type": verify_type
        }
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def click_verify(self, image, label_image=None, extra=None, verify_type="30100"):
        # 通用任意点选1~4个坐标 30009
        # 通用文字点选1(extra,点选文字逗号隔开,原图) 30100
        # 定制-文字点选2(extra="click",原图) 30103
        # 定制-单图文字点选 30102
        # 定制-图标点选1(原图) 30104
        # 定制-图标点选2(原图,extra="icon") 30105
        # 定制-语序点选1(原图,extra="phrase") 30106
        # 定制-语序点选2(原图) 30107
        # 定制-空间推理点选1(原图,extra="请点击xxx") 30109
        # 定制-空间推理点选1(原图,extra="请_点击_小尺寸绿色物体。") 30110
        # 定制-tx空间点选(extra="请点击侧对着你的字母") 50009
        # 定制-tt_空间点选 30101
        # 定制-推理拼图1(原图,extra="交换2个图块") 30108
        # 定制-xy4九宫格点选(原图,label_image,image) 30008
        payload = {
            "image": base64.b64encode(image).decode(),
            # "label_image": base64.b64encode(label_image).decode(),
            "token": self._token,
            "type": verify_type
        }
        if extra:
            payload['extra'] = extra
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def rotate(self, out_ring_image, inner_circle_image):
        # 定制-X度单图旋转  90007
        # payload = {
        #     "image": base64.b64encode(image).decode(),
        #     "token": self._token,
        #     "type": "90007"
        # }
        # 定制-Tt双图旋转,2张图,内圈图,外圈图  90004
        payload = {
            "out_ring_image": base64.b64encode(out_ring_image).decode(),
            "inner_circle_image": base64.b64encode(inner_circle_image).decode(),
            "token": self._token,
            "type": "90004"
        }
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def google_verify(self, googlekey, pageurl, invisible=1, data_s=""):
        _headers = {
            'Content-Type': 'application/json'
        }
        """
        第一步，创建验证码任务
        :param
        :return taskId : string 创建成功的任务ID
        """
        url = "http://122.9.52.147/api/YmServer/funnelApi"
        payload = json.dumps({
            "token": self._token,
            # "type": "40011", ## v3
            "type": "40010",  ## v2
            "googlekey": googlekey,
            "enterprise": 1,  ## 是否为企业版
            "pageurl": pageurl,
            "invisible": invisible,
            "data-s": data_s,
            # 'action':"TEMPLATE" #V3必传

        })
        # 发送JSON格式的数据
        result = requests.request("POST", url, headers=_headers, data=payload).json()
        print(result)
        # {'msg': '识别成功', 'code': 10000, 'data': {'code': 0, 'captchaId': '51436618130', 'recordId': '74892'}}
        captcha_id = result.get('data').get("captchaId")
        record_id = result.get('data').get("recordId")
        times = 0
        while times < 150:
            try:
                url = f"http://122.9.52.147/api/YmServer/funnelApiResult"
                data = {
                    "token": self._token,
                    "captchaId": captcha_id,
                    "recordId": record_id
                }
                result = requests.post(url, headers=_headers, json=data).json()
                print(result)
                # {'msg': '结果准备中，请稍后再试', 'code': 10009, 'data': []}
                if result['msg'] == "结果准备中，请稍后再试":
                    time.sleep(5)
                    times += 5
                    continue
                if result['msg'] == '请求成功' and result['code'] == 10001:
                    print(result['data']['data'])
                    return result['data']['data']
                    # {'msg': '请求成功', 'code': 10001, 'data': {'data': '03AGdBq2611GTOgA2v9HUpMMEUE70p6dwOtYyHJQK4xhdKF0Y8ouSGsFZt647SpJvZ22qinYrm6MYBJGFQxMUIApFfSBN6WTGspk6DmFdQAoWxynObRGV7qNMQOjZ_m4w3_6iRu8SJ3vSUXH_HHuA7wXARJbKEpU4J4R921NfpKdahgeFD8rK1CFYAqLd5fz4l-8_VRmRE83dRSfkgyTN338evQ1doWKJRipZbk4ie-89Ud0KGdOsP4QzG3stRZgj2oaEoMDSAP62vxKGYqtDEqTcwtlgo-ot3rF5SmntaoKGwcKPo0NrekWA5gtj0vqKLU6lY2GcnSci_tgBzBwuH40uvyR1PFu02VK_E44mopJ7FOO4cUukNaLGqypU2YCA8QuaaebOIoCMU7RGqGs_41RYNCG1GSdthiwcwk2hHFbi-TXuICXSwh4Er5mgVW9A3t_9Ndp0eJcyr3HtuJrcA7BtlcgruuQxK5h4Ew4ert4KPH_aQGN9ww5VsUtbSManzUDnUOs7aEdvFk1DOOPmLys-aX20ZFN2CcQcZZSO-7HZpZZt3EDeWWE5S02HFDY8gl3_0xqIts8774Tr4GMVJaddG0NR6pcBFC11FqNcK2a18gM3gaKDy3_2ZMeSU4nj4NWwoAhPjQN2BS8JxX4kKVpX4rD959kc93vczVD3TYD6_4GJahGSpBvM7Y5_GGIdLL8imXde1R35mZnEcFYXQ40zcy3DdJFkk_gzGTVOEb1Q1IZpjMxzCxyGgwjgL9dtDIgst5H5CSZoerX_Lz-DmsBvYIYZdpbPLEMROx9MODImaEw8Cp6M8Xj7_foijiGE9hh-pzJSTlKl3HytiSUyJJ7r1BssrX5C_TFWxl0IXNg8azP8H-ZIOWwnYlMWCS1w9piHdoLg5zACiYIN3Txdlsvi61MuPmzJggJd1_dlyMdAlzb5_zdfweqj0_Ko1ODP378YT7sV7LECgRj5QJU6sF5nlf4m2g5sFypBw9GFAkEE-OaWGYxRJOy2ioU41ggAJIkcza2B_N5AL2KLROtm0-c2MxplM4ZzHxrUv9A24zlgzo3Pz4NONwU_gaOcDB7j1dZKXD8UaoIrZv0BTd8JeojYowm9Usdg7Rt4Fpo_vDLJdrEUfbxVlXieDD9Fr1fu72-d4AduT_J3n-rIhyX4gFav-KfP-qOxqOZsmjXZirsBxZs7042NYeirRYnLv35cxIAJARz03FJmeKViUivwC5mCWw64hjRad9XyyBOP2n8KFOrTXhPskC-WwEfksGtfLxi6VW76FHGvRdwHXzMwVfNqe3P5H_WZUc-vxeTAsTnqZz3WA97lM4MLrX0nTZYgXxCEiS6raSOiEMqcx_Nv7Zxre-abj4LZRbFpH8nx1SEiaOV2Dm-a1iPFEmCs0L4kDtt6VImSVIQaTOAd3KFSo7W_XTvRPsQJOtblrcKyuagztX_Yr0lT0YqN9I9MZAARo7M5OfwSLJW16rdmp4NuRefEvNPNHO2cVh1Xha1qNGuF_QDvWFFmWG0Y6IbRqLmF-Dv8BY4TWyOeVnADJftGQw2QSr8RmbCHryA'}}
            except Exception as e:
                print(e)
                continue

    def fun_captcha_verify(self, publickey, pageurl, verify_type="40007"):
        # 定制类接口-Hcaptcha 40007
        payload = {
            "publickey": publickey,
            "pageurl": pageurl,
            "token": self._token,
            "type": verify_type
        }
        resp = requests.post(self._custom_url, headers=self._headers, data=json.dumps(payload))
        print(resp.text)
        return resp.json()['data']['data']

    def hcaptcha_verify(self):
        # 定制类接口-Hcaptcha
        _headers = {
            'Content-Type': 'application/json'
        }
        _custom_url = "http://api.jfbym.com/api/YmServer/funnelApi"
        payload = {
            "sitekey": "",
            "pageurl": "",
            "token": self._token,
            "type": '50013'
        }
        result = requests.post(_custom_url, headers=_headers, data=json.dumps(payload)).json()
        print(result)
        captcha_id = result.get('data').get("captchaId")
        record_id = result.get('data').get("recordId")
        times = 0
        while times < 150:
            try:
                url = f"http://api.jfbym.com/api/YmServer/funnelApiResult"
                data = {
                    "token": self._token,
                    "captchaId": captcha_id,
                    "recordId": record_id
                }
                result = requests.post(url, headers=_headers, json=data).json()
                print(result)
                # {'msg': '结果准备中，请稍后再试', 'code': 10009, 'data': []}
                if result['msg'] == "结果准备中，请稍后再试":
                    time.sleep(5)
                    times += 5
                    continue
                if result['msg'] == '请求成功' and result['code'] == 10001:
                    print(result['data']['data'])
                    return result['data']['data']
                    # {'msg': '请求成功', 'code': 10001, 'data': {'data': '03AGdBq2611GTOgA2v9HUpMMEUE70p6dwOtYyHJQK4xhdKF0Y8ouSGsFZt647SpJvZ22qinYrm6MYBJGFQxMUIApFfSBN6WTGspk6DmFdQAoWxynObRGV7qNMQOjZ_m4w3_6iRu8SJ3vSUXH_HHuA7wXARJbKEpU4J4R921NfpKdahgeFD8rK1CFYAqLd5fz4l-8_VRmRE83dRSfkgyTN338evQ1doWKJRipZbk4ie-89Ud0KGdOsP4QzG3stRZgj2oaEoMDSAP62vxKGYqtDEqTcwtlgo-ot3rF5SmntaoKGwcKPo0NrekWA5gtj0vqKLU6lY2GcnSci_tgBzBwuH40uvyR1PFu02VK_E44mopJ7FOO4cUukNaLGqypU2YCA8QuaaebOIoCMU7RGqGs_41RYNCG1GSdthiwcwk2hHFbi-TXuICXSwh4Er5mgVW9A3t_9Ndp0eJcyr3HtuJrcA7BtlcgruuQxK5h4Ew4ert4KPH_aQGN9ww5VsUtbSManzUDnUOs7aEdvFk1DOOPmLys-aX20ZFN2CcQcZZSO-7HZpZZt3EDeWWE5S02HFDY8gl3_0xqIts8774Tr4GMVJaddG0NR6pcBFC11FqNcK2a18gM3gaKDy3_2ZMeSU4nj4NWwoAhPjQN2BS8JxX4kKVpX4rD959kc93vczVD3TYD6_4GJahGSpBvM7Y5_GGIdLL8imXde1R35mZnEcFYXQ40zcy3DdJFkk_gzGTVOEb1Q1IZpjMxzCxyGgwjgL9dtDIgst5H5CSZoerX_Lz-DmsBvYIYZdpbPLEMROx9MODImaEw8Cp6M8Xj7_foijiGE9hh-pzJSTlKl3HytiSUyJJ7r1BssrX5C_TFWxl0IXNg8azP8H-ZIOWwnYlMWCS1w9piHdoLg5zACiYIN3Txdlsvi61MuPmzJggJd1_dlyMdAlzb5_zdfweqj0_Ko1ODP378YT7sV7LECgRj5QJU6sF5nlf4m2g5sFypBw9GFAkEE-OaWGYxRJOy2ioU41ggAJIkcza2B_N5AL2KLROtm0-c2MxplM4ZzHxrUv9A24zlgzo3Pz4NONwU_gaOcDB7j1dZKXD8UaoIrZv0BTd8JeojYowm9Usdg7Rt4Fpo_vDLJdrEUfbxVlXieDD9Fr1fu72-d4AduT_J3n-rIhyX4gFav-KfP-qOxqOZsmjXZirsBxZs7042NYeirRYnLv35cxIAJARz03FJmeKViUivwC5mCWw64hjRad9XyyBOP2n8KFOrTXhPskC-WwEfksGtfLxi6VW76FHGvRdwHXzMwVfNqe3P5H_WZUc-vxeTAsTnqZz3WA97lM4MLrX0nTZYgXxCEiS6raSOiEMqcx_Nv7Zxre-abj4LZRbFpH8nx1SEiaOV2Dm-a1iPFEmCs0L4kDtt6VImSVIQaTOAd3KFSo7W_XTvRPsQJOtblrcKyuagztX_Yr0lT0YqN9I9MZAARo7M5OfwSLJW16rdmp4NuRefEvNPNHO2cVh1Xha1qNGuF_QDvWFFmWG0Y6IbRqLmF-Dv8BY4TWyOeVnADJftGQw2QSr8RmbCHryA'}}
            except Exception as e:
                print(e)
                continue


if __name__ == '__main__':
    y = YdmVerify()
    y.hcaptcha_verify()
