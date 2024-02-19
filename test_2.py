import time
import unittest

import pyttsx4
import winsound


class MyTestCase(unittest.TestCase):
    def test_something(self):
        self.assertEqual(self.sound_signal(32767, 500), None)

    @staticmethod
    def system_prompt_tone(sound_type) -> None:
        """系统提示音
        :param sound_type: 提示音类型(1:警告, 2:错误, 3:询问, 4:信息, 5:系统启动, 6:系统关闭)"""
        if sound_type == '系统警告':
            winsound.PlaySound('SystemAsterisk', winsound.SND_ALIAS)
        elif sound_type == '系统错误':
            winsound.PlaySound('SystemExclamation', winsound.SND_ALIAS)
        elif sound_type == '系统询问':
            winsound.PlaySound('SystemQuestion', winsound.SND_ALIAS)
        elif sound_type == '系统信息':
            winsound.PlaySound('SystemHand', winsound.SND_ALIAS)
        elif sound_type == '系统启动':
            winsound.PlaySound('SystemStart', winsound.SND_ALIAS)
        elif sound_type == '系统关闭':
            winsound.PlaySound('SystemExit', winsound.SND_ALIAS)

    @staticmethod
    def sound_signal(frequency: int, duration: int, times: int = 1, interval: int = 0) -> None:
        """播放音频信号
        :param frequency: 频率(37~32767)
        :param duration: 持续时间(毫秒)
        :param times: 次数
        :param interval: 间隔时间(毫秒)"""
        try:
            for _ in range(times):
                winsound.Beep(frequency, duration)
                if interval:
                    time.sleep(interval / 1000)
        except RuntimeError:
            pass

    @staticmethod
    def play_audio(info: str, rate: int = 200) -> None:
        """播放TTS提示音"""
        try:
            engine = pyttsx4.init()
            engine.setProperty('rate', rate)  # 设置语速
            engine.say(info)
            engine.runAndWait()
        except Exception as e:
            print(e)


if __name__ == '__main__':
    unittest.main()
