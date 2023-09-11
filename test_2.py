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