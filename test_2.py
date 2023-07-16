import pandas as pd
from selenium import webdriver

driver = webdriver.Chrome()
driver.get("网页URL")

# 定位到表格
table = driver.find_element_by_id("table-id")

# 获取表格的HTML代码
html = table.get_attribute("outerHTML")

# 用pandas读取解析表格
df = pd.read_html(html)[0]

# 查看结果
print(df)