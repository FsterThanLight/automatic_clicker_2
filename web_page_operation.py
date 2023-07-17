import time

import openpyxl
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd


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
        # 保存的表格数据
        self.excel_path = None
        self.sheet_name = None

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
        # print('self.driver: ', self.driver)
        if self.driver is not None:
            self.driver.quit()

    def lookup_element(self, element_type_, timeout_type_):
        """查找元素
        :param element_type_: 元素类型
        :param timeout_type_: 超时错误"""

        def lookup_element_x(element_type__):
            """查找元素"""
            print('查找元素。')
            print(element_type__)
            if element_type__ == '元素ID':
                self.wait_for_action_element = self.driver.find_element(By.ID, self.element_wait_for_action)
            elif element_type__ == '元素名称':
                self.wait_for_action_element = self.driver.find_element(By.NAME, self.element_wait_for_action)
            elif element_type__ == 'xpath定位':
                self.wait_for_action_element = self.driver.find_element(By.XPATH, self.element_wait_for_action)

        try:
            lookup_element_x(element_type_)
        except NoSuchElementException:
            if timeout_type_ == '找不到元素自动跳过':
                pass
            else:
                time_wait = int(timeout_type_)
                # 继续查找元素，直到超时
                while time_wait > 0:
                    try:
                        lookup_element_x(element_type_)
                        break
                    except NoSuchElementException:
                        print('查找元素失败，正在重试。剩余' + str(time_wait) + '秒。')
                        # QApplication.processEvents()
                        # self.main_window_.plainTextEdit.appendPlainText('查找元素失败，正在重试。剩余' + str(time_wait) + '秒。')
                        time.sleep(1)
                        time_wait -= 1
                raise TimeoutException

    def perform_mouse_action(self, element_type_, timeout_type_, action, text=None):
        """鼠标操作
        :param action: 鼠标操作
        :param element_type_: 元素类型
        :param timeout_type_: 超时错误
        :param text: 输入内容"""
        self.chains = ActionChains(self.driver)
        # 查找元素(元素类型、超时错误)
        self.lookup_element(element_type_, timeout_type_)
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
            elif action == '读取网页表格':
                table_html = self.wait_for_action_element.get_attribute('outerHTML')
                df = pd.read_html(table_html)
                with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a') as writer:
                    df[0].to_excel(writer, index=False, sheet_name=self.sheet_name)

    def single_shot_operation(self, url, action, element_type_, element_value_, timeout_type_, text=None):
        """单步骤操作
        :param url: 网址
        :param action: 鼠标操作（左键单击、左键双击、右键单击、输入内容、读取网页表格）
        :param element_type_: 元素类型（元素ID、元素名称、xpath定位）
        :param element_value_: 元素值
        :param timeout_type_: 超时错误（找不到元素自动跳过、秒数）
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
        if action == '' or action is None:
            print('没有鼠标操作。')
            pass
        else:
            print('执行鼠标操作。')
            # 确定等待操作的元素
            self.element_wait_for_action = element_value_
            # 执行鼠标操作
            self.perform_mouse_action(action=action,
                                      element_type_=element_type_,
                                      timeout_type_=timeout_type_,
                                      text=text)

    def switch_to_frame(self, iframe_type, iframe_value, switch_type):
        """切换frame
        :param iframe_type: iframe类型（id或名称、xpath定位）
        :param iframe_value: iframe值
        :param switch_type: 切换类型（切换到指定frame，切换到上一级或切换到主文档）"""
        if switch_type == '切换到指定frame':
            if iframe_type == 'frame名称或ID：':
                self.driver.switch_to.frame(iframe_value)
            elif iframe_type == 'Xpath定位：':
                self.driver.switch_to.frame(self.driver.find_element(By.XPATH, iframe_value))
        elif switch_type == '切换到上一级文档':
            self.driver.switch_to.parent_frame()
        elif switch_type == '切换回主文档':
            self.driver.switch_to.default_content()


# WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
# driver.switch_to.parent_frame()
# waitElement.wait_element(driver, "XPATH", xpath, "ddd "):

if __name__ == '__main__':
    # 初始化功能类
    web = WebOption()

    # web.single_shot_operation(url='www.baidu.com',
    #                           action='',
    #                           element_value_='',
    #                           element_type_='',
    #                           text='',
    #                           timeout_type_=3)
    # #
    # # web.single_shot_operation(url='',
    # #                           action='输入内容',
    # #                           element_value_='kw',
    # #                           element_type_='元素ID',
    # #                           text='python',
    # #                           timeout_type_=3)
    # element_value = 'kw'
    # element_type = '元素ID'
    # timeout_type = 3
    # cell_value = '德国'
    #
    # web.single_shot_operation(url='',
    #                           action='输入内容',
    #                           element_value_=element_value,
    #                           element_type_=element_type,
    #                           text=cell_value,
    #                           timeout_type_=timeout_type)

    web.excel_path = r'C:\Users\federalsadler\Desktop\1.xlsx'
    web.sheet_name = 'su'
    web.single_shot_operation(url='http://www.tianqihoubao.com/weather/top/chengdu.html',
                              action='读取网页表格',
                              element_value_='//*[@id="content"]/table',
                              element_type_='xpath定位',
                              text='',
                              timeout_type_=3)

    time.sleep(5)
    web.close_browser()
