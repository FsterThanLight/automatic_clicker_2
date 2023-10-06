import random
import time
import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class WebOption:
    def __init__(self, main_window=None, navigation=None):
        self.main_window = main_window
        self.navigation = navigation
        self.driver = None
        # 保存的表格数据
        self.excel_path = None
        self.sheet_name = None
        # 拖动元素的距离
        self.distance_x = 0
        self.distance_y = 0
        # 输入的文本
        self.text = None

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
            driver_ = webdriver.Chrome(service=service)
            driver_.quit()
        except ConnectionError:
            QMessageBox.warning(self.navigation, '警告', '驱动安装失败，请重试。', QMessageBox.Yes)

    def close_browser(self):
        """关闭浏览器驱动"""
        print('关闭浏览器驱动。')
        if self.driver is not None:
            self.driver.quit()

    def lookup_element(self, element_value_, element_type_, timeout_type_):
        """查找元素
        :param element_value_: 未开始查找的元素值
        :param element_type_: 元素类型
        :param timeout_type_: 超时错误"""
        # 查找元素(元素类型、超时错误)
        # 等待到指定元素出现
        try:
            print('正在查找元素' + element_value_)
            target_ele = None
            if element_type_ == 'xpath定位':
                target_ele = WebDriverWait(self.driver, timeout_type_).until(
                    EC.presence_of_element_located((By.XPATH, element_value_)))
            elif element_type_ == '元素名称':
                target_ele = WebDriverWait(self.driver, timeout_type_).until(
                    EC.presence_of_element_located((By.NAME, element_value_)))
            elif element_type_ == '元素ID':
                target_ele = WebDriverWait(self.driver, timeout_type_).until(
                    EC.presence_of_element_located((By.ID, element_value_)))
            return target_ele
        except TimeoutException:
            return None
        except NoSuchElementException:
            return None

    def switch_to_frame(self, iframe_type, iframe_value, switch_type):
        """切换frame
        :param iframe_type: iframe类型（id或名称：、xpath定位：）
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

    def switch_to_window(self, window_type, window_value):
        """切换窗口
        :param window_type: 窗口类型（窗口名称或ID：、窗口标题：）
        :param window_value: 窗口值"""
        if window_type == '窗口ID：':
            self.driver.switch_to.window(self.driver.window_handles[int(window_value)])
        elif window_type == '窗口标题：':
            for handle in self.driver.window_handles:
                self.driver.switch_to.window(handle)
                if self.driver.title in window_value:
                    break

    def perform_mouse_action(self, element_value_, element_type_, timeout_type_, action):
        """鼠标操作
        :param element_value_: 未开始查找的元素值
        :param action: 鼠标操作
        :param element_type_: 元素类型
        :param timeout_type_: 超时错误"""
        # 查找元素(元素类型、超时错误)
        target_ele = self.lookup_element(element_value_, element_type_, timeout_type_)
        if target_ele is not None:
            print('找到网页元素，执行鼠标操作。')
            # self.main_window_.plainTextEdit.appendPlainText('找到网页元素，执行鼠标操作。')
            # QApplication.processEvents()
            if action == '左键单击':
                ActionChains(self.driver).click(target_ele).perform()
            elif action == '左键双击':
                ActionChains(self.driver).double_click(target_ele).perform()
            elif action == '右键单击':
                ActionChains(self.driver).context_click(target_ele).perform()
            elif action == '输入内容':
                # 检查元素是否为可读属性
                read_only = target_ele.get_attribute('readonly')
                if not read_only:
                    target_ele.send_keys(self.text)
                else:
                    print('元素为只读属性，正在移除只读属性。')
                    self.driver.execute_script("arguments[0].removeAttribute('readonly');", target_ele)
                    target_ele.click()
                    target_ele.clear()
                    target_ele.send_keys(self.text)
            elif action == '读取网页表格':
                table_html = target_ele.get_attribute('outerHTML')
                df = pd.read_html(table_html)
                df1 = pd.DataFrame(df[0])
                df1.to_excel(self.excel_path, index=False, sheet_name=self.sheet_name)
                # with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a') as writer:
                #     df[0].to_excel(writer, index=False, sheet_name=self.sheet_name)
            elif action == '拖动元素':
                ActionChains(self.driver).click_and_hold(target_ele).perform()
                ActionChains(self.driver).move_by_offset(self.distance_x, self.distance_y +
                                                         random.randint(-5, 5)).perform()
                ActionChains(self.driver).release().perform()
        elif target_ele is None:
            raise TimeoutException

    def single_shot_operation(self, url, action, element_type_, element_value_, timeout_type_):
        """单步骤操作
        :param url: 网址
        :param action: 鼠标操作（左键单击、左键双击、右键单击、输入内容、读取网页表格、拖动元素）
        :param element_type_: 元素类型（元素ID、元素名称、xpath定位）
        :param element_value_: 元素值
        :param timeout_type_: 超时错误（找不到元素自动跳过、秒数）"""

        def open_url(url_):
            """打开网页或者直接跳过"""
            if url_ == '' or url_ is None:
                pass
            else:
                if url_[:7] != 'http://' and url_[:8] != 'https://':
                    url_ = 'http://' + url_

                chrome_options = webdriver.ChromeOptions()
                # 添加选项配置：  # 但是用程序打开的网页的window.navigator.webdriver仍然是true。
                chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
                chrome_options.add_experimental_option("detach", True)
                # # 去掉window.navigator.webdriver的特性
                # chrome_options.add_argument("disable-blink-features=AutomationControlled")
                # 设置为无头浏览器：不会显示出操作浏览器的过程
                # chrome_options.add_argument('--headless')
                # chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
                chrome_options.add_argument('--start-maximized')
                # 初始化浏览器并打开网页
                self.driver = webdriver.Chrome(options=chrome_options)
                self.driver.get(url_)
                # 窗口最大化
                # self.driver.maximize_window()

                time.sleep(1)

        open_url(url)
        if action == '' or action is None:
            print('没有鼠标操作。')
            pass
        else:
            print('执行鼠标操作。')
            # 执行鼠标操作
            self.perform_mouse_action(action=action,
                                      element_type_=element_type_,
                                      timeout_type_=timeout_type_,
                                      element_value_=element_value_)
