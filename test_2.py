import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

wd = webdriver.Chrome()
wd.get('http://www.baidu.com')

# wd是webdriver对象，10是最长等待时间，0.5是每0.5秒去查询对应的元素。until后面跟的等待具体条件，EC是判断条件，检查元素是否存在于页面的 DOM 上。
# 这行可以理解为 每0.5s连接到百度的首页看看，有没有出来
login_btn = WebDriverWait(wd, 10, 0.5).until(EC.presence_of_element_located((By.ID, "s-top-loginbtn")))

# 再举个例子 比如网速慢， 【百度一下】 左边的输入框没出来，那我们就设置，如果出现就输入 查找的关键字
WebDriverWait(wd, 10, 0.5).until(EC.presence_of_element_located((By.XPATH, "//input[@id='kw']")))
input_ = wd.find_element_by_xpath("//input[@id='kw']")
time.sleep(0.1)
input_.send_keys("python")

# 再举个例子 比如网速慢， 【百度一下】 这四个字没出来(按钮)，那我们就设置，如果出现就输入 查找的关键字，然后点击
WebDriverWait(wd, 10, 0.5).until(EC.presence_of_element_located((By.XPATH, "//input[@id='su']")))
baiduyixia = wd.find_element_by_xpath("//input[@id='su']")
time.sleep(0.1)
baiduyixia.click()

time.sleep(2)
wd.close()
