import sys

from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QApplication, QListView, QMainWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("ListView with Numbers")
        self.setGeometry(100, 100, 400, 300)

        # 创建一个 QListView
        self.listView = QListView(self)
        self.listView.setGeometry(10, 10, 380, 280)

        # 创建一个 QStandardItemModel
        self.model = QStandardItemModel()

        # 向模型中添加数据项
        for i in range(10):
            item = str(i + 1)  # 序号从1开始
            self.model.appendRow(QStandardItem(item))

        # 将模型设置到 ListView 中
        self.listView.setModel(self.model)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
