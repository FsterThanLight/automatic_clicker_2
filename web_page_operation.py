from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains
import time


class WebOption:
    def __init__(self):
        # 打开浏览器
        self.driver = webdriver.Chrome()
        # 元素id
        self.element_id = None
        # 鼠标操作
        self.wait_for_action_element = None
        self.chains = ActionChains(self.driver)

    @staticmethod
    def install_browser_driver():
        """安装谷歌浏览器的驱动"""
        service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.quit()

    def open_web(self, url):
        """打开指定网页"""
        self.driver.get(url)

    def lookup_element(self):
        """查找元素"""
        # 查找输入框，并输入内容
        self.wait_for_action_element = self.driver.find_element(By.ID, self.element_id)

    def perform_mouse_action(self, action, text=None):
        """鼠标操作"""
        self.lookup_element()
        if action == '左键单击':
            self.chains.click(self.wait_for_action_element).perform()
        elif action == '左键双击':
            self.chains.double_click(self.wait_for_action_element).perform()
        elif action == '右键单击':
            self.chains.context_click(self.wait_for_action_element).perform()
        elif action == '输入内容':
            self.wait_for_action_element.send_keys(text)


if __name__ == '__main__':
    web = WebOption()
    web.open_web('https://www.baidu.com')
    time.sleep(5)
    web.element_id = 'kw'
    web.perform_mouse_action('输入内容', 'python')
    time.sleep(2)
    web.element_id = 'su'
    web.perform_mouse_action('右键单击')
    time.sleep(5)
