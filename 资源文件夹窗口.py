import os

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QDialog, QFileDialog

from 数据库操作 import global_write_to_database, sqlitedb, close_database, extract_global_parameter
from 窗体.global_s import Ui_Global


class Global_s(QDialog, Ui_Global):
    """全局参数设置窗体"""

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setupUi(self)
        # 去除帮助按钮
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.refresh_listview()  # 刷新listview
        self.pushButton.clicked.connect(self.select_file)  # 添加图像文件夹路径
        self.pushButton_2.clicked.connect(self.delete_listview)  # 删除listview中的项

    def select_file(self):
        """打开选择文件窗口,并将路径写入数据库"""
        fil_path = QFileDialog.getExistingDirectory(self, "选择存储目标图像的文件夹")
        if fil_path != '':
            global_write_to_database(os.path.normpath(fil_path))
        self.refresh_listview()

    def delete_listview(self):
        """删除listview中选中的那行数据"""

        # 获取选中的行的值
        def delete_data(value_):
            """删除数据库中的数据"""
            # 连接数据库
            cursor, conn = sqlitedb()
            cursor.execute("DELETE FROM 全局参数 WHERE 资源文件夹路径 = ?", (value_,))  # 删除数据
            cursor.execute("DELETE FROM 全局参数 WHERE 资源文件夹路径 is NULL and 分支表名 is Null")  # 删除无用数据条
            conn.commit()
            close_database(cursor, conn)

        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text()
            delete_data(value)  # 删除数据库中的数据
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

        res_folder_path = extract_global_parameter('资源文件夹路径')  # 获取数据库中的数据
        add_listview(res_folder_path, self.listView)