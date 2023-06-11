import sys
import time

from PyQt5.QtWidgets import QMessageBox, QApplication
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


class WebOption:
    def __init__(self, main_window=None, navigation=None):
        self.main_window = main_window
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
    
    def close_browser(self):
        """关闭浏览器驱动"""
        print('关闭浏览器驱动。')
        print('self.driver: ', self.driver)
        if self.driver is not None:
            self.driver.quit()

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
                        # QApplication.processEvents()
                        # self.main_window.plainTextEdit.appendPlainText('查找元素失败，正在重试。剩余' + str(time_wait) + '秒。')
                        time.sleep(1)
                        time_wait -= 1
                raise TimeoutException

    def perform_mouse_action(self, action, element_type, timeout_type, text=None):
        """鼠标操作"""
        self.chains = ActionChains(self.driver)
        # 查找元素(元素类型、超时错误)
        self.lookup_element(element_type, timeout_type)

        if self.wait_for_action_element is not None:
            print('找到网页元素，执行鼠标操作。')
            # self.main_window.plainTextEdit.appendPlainText('找到网页元素，执行鼠标操作。')
            if action == '左键单击':
                self.chains.click(self.wait_for_action_element).perform()
            elif action == '左键双击':
                self.chains.double_click(self.wait_for_action_element).perform()
            elif action == '右键单击':
                self.chains.context_click(self.wait_for_action_element).perform()
            elif action == '输入内容':
                self.wait_for_action_element.send_keys(text)

    def single_shot_operation(self, url, action, element_type, element_value, timeout_type, text=None):
        """单步骤操作"""
        if url == '' or url is None:
            pass
        else:
            if url[:7] != 'http://' and url[:8] != 'https://':
                url = 'http://' + url
            # 初始化浏览器并打开网页
            self.driver = webdriver.Chrome()
            self.driver.get(url)
            time.sleep(1)

        if element_type == '元素ID':
            self.element_id = element_value
        elif element_type == '元素名称':
            self.element_name = element_value

        self.perform_mouse_action(action, element_type, timeout_type, text)


if __name__ == '__main__':
    # 初始化功能类
    web = WebOption()

    web.single_shot_operation(url='www.baidu.com',
                              action='输入内容',
                              element_value='wd',
                              element_type='元素名称',
                              text='python',
                              timeout_type='找不到元素自动跳过')

    web.single_shot_operation(url='',
                              action='左键单击',
                              element_value='su',
                              element_type='元素ID',
                              timeout_type='5')

    time.sleep(10)
    web.close_browser()
