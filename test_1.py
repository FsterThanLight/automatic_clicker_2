from pandas import ExcelFile



if __name__ == "__main__":
    excel_path = r"C:\Users\Administrator\Desktop\personal_data\数据记录.xlsx"
    xl = ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    print(sheet_names)