import time

from PyQt5.QtWidgets import QMessageBox
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
        # 等待操作的元素
        self.element_wait_for_action = None
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
        """查找元素
        :param element_type: 元素类型
        :param timeout_type: 超时错误"""

        def lookup_element_x(element_type_x):
            """查找元素"""
            if element_type_x == '元素ID':
                self.wait_for_action_element = self.driver.find_element(By.ID, self.element_wait_for_action)
            elif element_type_x == '元素名称':
                self.wait_for_action_element = self.driver.find_element(By.NAME, self.element_wait_for_action)
            elif element_type_x == '元素类名':
                self.wait_for_action_element = self.driver.find_element(By.XPATH, self.element_wait_for_action)

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
                        # self.main_window_.plainTextEdit.appendPlainText('查找元素失败，正在重试。剩余' + str(time_wait) + '秒。')
                        time.sleep(1)
                        time_wait -= 1
                raise TimeoutException

    def perform_mouse_action(self, action, element_type, timeout_type, text=None):
        """鼠标操作
        :param action: 鼠标操作
        :param element_type: 元素类型
        :param timeout_type: 超时错误
        :param text: 输入内容"""
        self.chains = ActionChains(self.driver)
        # 查找元素(元素类型、超时错误)
        self.lookup_element(element_type, timeout_type)

        if self.wait_for_action_element is not None:
            print('找到网页元素，执行鼠标操作。')
            # self.main_window_.plainTextEdit.appendPlainText('找到网页元素，执行鼠标操作。')
            if action == '左键单击':
                self.chains.click(self.wait_for_action_element).perform()
            elif action == '左键双击':
                self.chains.double_click(self.wait_for_action_element).perform()
            elif action == '右键单击':
                self.chains.context_click(self.wait_for_action_element).perform()
            elif action == '输入内容':
                self.wait_for_action_element.send_keys(text)

    def single_shot_operation(self, url, action, element_type, element_value, timeout_type, text=None):
        """单步骤操作
        :param url: 网址
        :param action: 鼠标操作
        :param element_type: 元素类型
        :param element_value: 元素值
        :param timeout_type: 超时错误
        :param text: 输入内容"""

        def open_url(url_):
            """打开网页或者直接跳过"""
            if url_ == '' or url_ is None:
                pass
            else:
                if url_[:7] != 'http://' and url_[:8] != 'https://':
                    url_ = 'http://' + url_
                # 初始化浏览器并打开网页
                self.driver = webdriver.Chrome()
                self.driver.get(url_)
                time.sleep(1)

        open_url(url)
        # 确定等待操作的元素
        self.element_wait_for_action = element_value
        # 执行鼠标操作
        self.perform_mouse_action(action, element_type, timeout_type, text)


# WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
# driver.switch_to.parent_frame()
# waitElement.wait_element(driver, "XPATH", xpath, "ddd "):

if __name__ == '__main__':
    # 初始化功能类
    web = WebOption()

    web.single_shot_operation(url='www.baidu.com',
                              action='输入内容',
                              element_value='//*[@id="kw"]',
                              element_type='元素类名',
                              text='python',
                              timeout_type=3)

    time.sleep(10)
    web.close_browser()
