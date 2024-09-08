import os

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from ini控制 import set_window_size, save_window_size, extract_resource_folder_path, del_resource_folder_path, \
    writes_to_resource_folder_path, move_resource_folder_up_and_down
from 窗体.global_s import Ui_Global


class Global_s(QDialog, Ui_Global):
    """全局参数设置窗体"""

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setupUi(self)
        # 去除帮助按钮
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        set_window_size(self)  # 获取上次退出时的窗口大小
        # 绑定事件
        self.refresh_listview()  # 刷新listview
        self.pushButton.clicked.connect(self.select_file)  # 添加图像文件夹路径
        self.pushButton_2.clicked.connect(self.delete_listview)  # 删除listview中的项
        self.listView.doubleClicked.connect(self.open_select_listview)  # 双击打开listview中的文件夹路径
        # 上下移动选中的路径
        self.pushButton_4.clicked.connect(lambda: self.move_up_down('up'))
        self.pushButton_5.clicked.connect(lambda: self.move_up_down('down'))

    def select_file(self):
        """打开选择文件窗口,并将路径写入数据库"""
        fil_path = QFileDialog.getExistingDirectory(
            parent=self,
            caption="选择存储目标图像的文件夹",
            directory=os.path.expanduser("~")
        )
        if fil_path != '':
            # 检查路径中是否有中文
            if any('\u4e00' <= char <= '\u9fff' for char in fil_path):
                QMessageBox.critical(self, '警告', '资源文件夹路径中暂不允许含有中文字符，请重新选择！')
                return
            writes_to_resource_folder_path(os.path.normpath(fil_path))
        self.refresh_listview()

    def open_select_listview(self):
        """打开选中的listview中的文件夹路径"""
        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text()
            os.startfile(value)
        except Exception as e:
            # 删除不存在的文件夹路径
            print(e)
            self.delete_listview()
            QMessageBox.critical(self, '错误', '该文件夹路径不存在！已从列表中删除！')

    def delete_listview(self):
        """删除listview中选中的那行数据"""
        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text()
            del_resource_folder_path(value)  # 删除数据库中的数据
            self.refresh_listview()  # 刷新listview
        except Exception as e:
            print(e)

    def refresh_listview(self):
        """刷新listview"""

        def add_listview(list_, listview):
            """添加listview"""
            model = QStandardItemModel()
            listview.setModel(model)
            for item in list_:
                model.appendRow(QStandardItem(item))

        res_folder_path = extract_resource_folder_path()  # 获取数据库中的数据
        add_listview(res_folder_path, self.listView)

    def move_up_down(self, direction):
        """上移或下移"""
        indexes = self.listView.selectedIndexes()
        if not indexes:
            return
        path = indexes[0].data()
        move_resource_folder_up_and_down(path, direction)
        self.refresh_listview()
        # 选中移动后的项
        for i in range(self.listView.model().rowCount()):
            if self.listView.model().item(i).text() == path:
                self.listView.setCurrentIndex(self.listView.model().index(i, 0))
                break

    def closeEvent(self, event):
        """关闭窗口时触发"""
        # 窗口大小
        save_window_size(self.width(), self.height(), self.windowTitle())
