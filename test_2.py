import time

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
driver.get('https://www.baidu.com/')
driver.maximize_window()
driver.implicitly_wait(5)
# # 定位设置元素
# set_ele = driver.find_element(By.XPATH, '//*[@id="su"]')
# # 第一步：创建一个鼠标操作的对象
# # ActionChains(driver).move_to_element_with_offset(to_element=set_ele, xoffset=random.randint(30, 35),
# #                                                  yoffset=random.randint(30, 32)).perform()
# ActionChains(driver).context_click(on_element=set_ele).perform()
# time.sleep(5)


# 获取当前窗口信息及当前url
current_window = driver.current_window_handle
print("当前窗口信息:", current_window)

current_url = driver.current_url
print("当前窗口url:", current_url)
# 获取浏览器全部窗口句柄
handles = driver.window_handles

print("获取浏览器全部窗口句柄:", handles)
# 切换到新的窗口

driver.switch_to.window(handles[1])

current_url = driver.current_url
print("当前窗口url:", current_url)