import openpyxl
import pandas as pd

# 创建一个ExcelWriter对象
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')

# 读取已有的Excel文件
book = openpyxl.load_workbook('output.xlsx')

# 将已有的sheet页添加到ExcelWriter对象中
writer.book = book

# 将DataFrame写入新的sheet页中
df.to_excel(writer, sheet_name='new_sheet', index=False)

# 保存Excel文件
writer.save()
