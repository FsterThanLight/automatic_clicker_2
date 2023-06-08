import time

from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


class WebOption:
    def __init__(self, navigation=None):
        self.navigation = navigation
        self.driver = None
        # 元素id和名称
        self.element_id = None
        self.element_name = None
        # 鼠标操作
        self.wait_for_action_element = None
        self.chains = None

    def web_open_test(self, url):
        """打开网页"""
        if url == '':
            url = 'https://www.cn.bing.com/'
        self.driver = webdriver.Chrome()
        try:
            self.driver.get(url)
            time.sleep(1)
            self.driver.quit()
            time.sleep(1)
            QMessageBox.information(self.navigation, '提示', '成功打开网页。', QMessageBox.Yes)
        except Exception as e:
            # 弹出错误提示
            print(e)
            QMessageBox.warning(self.navigation, '警告', '连接失败，请重试。',
                                QMessageBox.Yes)

    @staticmethod
    def install_browser_driver():
        """安装谷歌浏览器的驱动"""
        service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.quit()

    def lookup_element(self, element_type):
        """查找元素"""
        # 查找输入框，并输入内容
        if element_type == 'id':
            self.wait_for_action_element = self.driver.find_element(By.ID, self.element_id)
        elif element_type == 'name':
            self.wait_for_action_element = self.driver.find_element(By.NAME, self.element_name)

    def perform_mouse_action(self, action, element_type, text=None):
        """鼠标操作"""
        self.driver = webdriver.Chrome()
        self.chains = ActionChains(self.driver)
        self.lookup_element(element_type)
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
    web.driver = webdriver.Chrome()
    web.driver.get('https://www.baidu.com')
    time.sleep(3)
    web.element_id = 'kw'
    web.perform_mouse_action('输入内容', 'id', 'python')
    time.sleep(2)
    web.element_id = 'su'
    web.perform_mouse_action('右键单击', 'id')
    time.sleep(5)
