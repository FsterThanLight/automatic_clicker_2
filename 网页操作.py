import random
import time

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class WebOption:
    def __init__(self, outputmessage=None):
        self.out_mes = outputmessage
        # self.navigation = navigation
        self.driver = None
        # 保存的表格数据
        self.excel_path = None
        self.sheet_name = None
        # 拖动元素的距离
        self.distance_x = 0
        self.distance_y = 0
        # 输入的文本
        self.text = None
        # 是否测试
        self.is_test = False

    def web_open_test(self, url):
        """打开网页
        :param url: 网址
        :return: 打开网页是否成功，成功返回True，失败返回False。"""
        url = 'https://www.cn.bing.com/' if url == '' else \
            ('http://' + url) if not url.startswith(('http://', 'https://')) else url

        try:
            self.driver = webdriver.Chrome()
            self.driver.get(url)
            time.sleep(1)
            return True, '打开网页成功。'
        except Exception as e:
            # 弹出错误提示
            print(e)
            return False, str(type(e))

    def output_message(self, mes):
        """输出信息
        :param mes: 信息内容"""
        if self.out_mes is not None:
            self.out_mes.out_mes(mes, self.is_test)

    def install_browser_driver(self):
        """安装谷歌浏览器的驱动"""
        try:
            self.out_mes.out_mes('正在安装谷歌浏览器驱动...等待中...', True)
            service = ChromeService(executable_path=ChromeDriverManager().install())
            driver_ = webdriver.Chrome(service=service)
            driver_.quit()
            self.out_mes.out_mes('浏览器驱动安装成功。', True)
        except Exception as e:
            self.out_mes.out_mes(f'浏览器驱动安装失败，请重试，错误信息：{e}。', True)

    def close_browser(self):
        """关闭浏览器驱动"""
        print(f'关闭浏览器驱动：{self.driver}')
        if self.driver is not None:
            self.driver.quit()

    def lookup_element(self, element_value_, element_type_, timeout_type_):
        """查找元素
        :param element_value_: 未开始查找的元素值
        :param element_type_: 元素类型
        :param timeout_type_: 超时错误"""

        # 查找元素(元素类型、超时错误)
        # 等待到指定元素出现
        def get_locator(element_type):
            """根据元素类型获取对应的定位器"""
            locators = {
                'xpath定位': By.XPATH,
                '元素名称': By.NAME,
                '元素ID': By.ID,
                # 添加其他可能的定位方式
            }
            return locators.get(element_type, By.XPATH)

        try:
            self.output_message(f'正在查找元素{element_value_}。')
            locator = get_locator(element_type_)  # 获取定位器
            target_ele = WebDriverWait(self.driver, timeout_type_).until(
                EC.presence_of_element_located((locator, element_value_))
            )
            return target_ele
        except Exception as e:
            print(e)
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
            self.output_message('找到网页元素，执行鼠标操作。')
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
                    # print('元素为只读属性，正在移除只读属性。')
                    self.output_message('元素为只读属性，正在移除只读属性。')
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

    def single_shot_operation(self, action, element_type_, element_value_, timeout_type_):
        """单步骤操作
        :param action: 鼠标操作（左键单击、左键双击、右键单击、输入内容、读取网页表格、拖动元素）
        :param element_type_: 元素类型（元素ID、元素名称、xpath定位）
        :param element_value_: 元素值
        :param timeout_type_: 超时错误（找不到元素自动跳过、秒数）"""
        if not action:
            print('没有鼠标操作。')
            return
        print('执行鼠标操作。')
        # 执行鼠标操作
        self.perform_mouse_action(action=action,
                                  element_type_=element_type_,
                                  timeout_type_=timeout_type_,
                                  element_value_=element_value_)

    def open_driver(self, url: str, judge: bool = True):
        """打开浏览器，返回浏览器驱动
        :param judge: 是否返回浏览器驱动
        :param url: 网址"""
        chrome_options = webdriver.ChromeOptions()
        # 添加选项配置：  # 但是用程序打开的网页的window.navigator.webdriver仍然是true。
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument('--start-maximized')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()
        self.driver.get(url if not url or url.startswith(('http://', 'https://')) else 'http://' + url)
        time.sleep(1)
        return self.driver if judge else None


if __name__ == '__main__':
    web_option = WebOption()
    web_option.web_open_test(url='https://www.baidu.com/')
    # XXX = web_option.open_driver('https://www.baidu.com/', True)
    #
    # xxx_option = WebOption()
    # xxx_option.driver = XXX
    # xxx_option.text = '百度'
    # xxx_option.single_shot_operation(action='输入内容',
    #                                  element_type_='xpath定位',
    #                                  element_value_='//*[@id="kw"]',
    #                                  timeout_type_=15)
