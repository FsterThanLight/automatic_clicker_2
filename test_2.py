import win32ui


def selectFile(path, type_):
    """选择文件
    :param path: 初始路径
    :param type_: 文件类型"""
    fspec = None
    if type_ == "exe":
        fspec = "执行文件 (*.exe, *.bat)|*.exe;*.bat||"
    elif type_ == "image":
        fspec = "图像文件 (*.jpg, *.jpeg, *.bmp, *.png)|*.jpg; *.jpeg; *.bmp; *.png||"
    dlg = win32ui.CreateFileDialog(2, None, None, 1, fspec, None)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir(path)  # 设置打开文件对话框中的初始显示目录
    flag = dlg.DoModal()
    if flag == 2:
        return None
    filename = dlg.GetPathName()  # 获取选择的文件名称
    return filename


def selectFolder(path):
    """选择文件夹
    :param path: 初始路径"""
    dlg = win32ui.CreateFileDialog(1, None, None, 0, None, None)  # 0表示打开文件夹对话框
    dlg.SetOFNInitialDir(path)  # 设置打开文件夹对话框中的初始显示目录
    dlg.DoModal()
    folder_path = dlg.GetPathName()  # 获取选择的文件夹路径
    return folder_path


# 示例用法
# initial_path = "C:\\Users\\YourUsername"
# selected_folder = selectFolder(initial_path)
# print("Selected Folder:", selected_folder)

if __name__ == "__main__":
    # filename = selectFile("F:", "image")
    # print(filename)
    selected_folder = selectFolder("E:")
    print("Selected Folder:", selected_folder)
