import time

from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.common import NoSuchElementException
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
        else:
            if url[:7] != 'http://' and url[:8] != 'https://':
                url = 'http://' + url

        self.driver = webdriver.Chrome()
        try:
            self.driver.get(url)
            time.sleep(1)
            self.driver.quit()
            QMessageBox.information(self.navigation, '提示', '连接成功。', QMessageBox.Yes)
        except Exception as e:
            # 弹出错误提示
            print(e)
            QMessageBox.warning(self.navigation, '警告', '连接失败，请重试。系统故障、网络故障或网址错误。',
                                QMessageBox.Yes)

    def install_browser_driver(self):
        """安装谷歌浏览器的驱动"""
        try:
            service = ChromeService(executable_path=ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)
            driver.quit()
        except ConnectionError:
            QMessageBox.warning(self.navigation, '警告', '驱动安装失败，请重试。', QMessageBox.Yes)

    def lookup_element(self, element_type, timeout_type):
        """查找元素"""
        def lookup_element_x(element_type_x):
            """查找元素"""
            if element_type_x == '元素ID':
                self.wait_for_action_element = self.driver.find_element(By.ID, self.element_id)
            elif element_type_x == '元素名称':
                self.wait_for_action_element = self.driver.find_element(By.NAME, self.element_name)
        try:
            lookup_element_x(element_type)
        except NoSuchElementException:
            if timeout_type == '找不到元素自动跳过':
                pass
            else:
                time_wait = int(timeout_type)
                # 继续查找元素，直到超时
                while time_wait > 0:
                    try:
                        lookup_element_x(element_type)
                        break
                    except NoSuchElementException:
                        print('查找元素失败，正在重试。剩余' + str(time_wait) + '秒。')
                        time.sleep(1)
                        time_wait -= 1

    def perform_mouse_action(self, action, element_type, timeout_type, text=None):
        """鼠标操作"""
        self.chains = ActionChains(self.driver)
        # 查找元素(元素类型、超时错误)
        self.lookup_element(element_type, timeout_type)

        if self.wait_for_action_element is not None:
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
    time.sleep(1)

    web.element_id = 'ss'
    web.perform_mouse_action('输入内容', '元素ID', '6', 'python')

    time.sleep(2)

    web.element_id = 'su'
    web.perform_mouse_action('右键单击', '元素ID', '找不到元素自动跳过')

    time.sleep(5)
