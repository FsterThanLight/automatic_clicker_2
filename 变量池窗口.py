import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QApplication, QTableView, QMainWindow, QMenu, QAction

from 数据库操作 import get_value_from_variable_table, set_value_to_variable_table


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("变量池")
        self.tableView = QTableView(self)  # 创建一个 QTableView
        # 设置表格风格
        self.tableView.horizontalHeader().setStretchLastSection(True)  # 设置最后一列拉伸至最大
        self.tableView.setAlternatingRowColors(True)  # 交替颜色
        self.tableView.setStyleSheet("QHeaderView::section{background:red;}")  # 设置表头背景色
        # 设置标题字体粗体
        self.tableView.horizontalHeader().setStyleSheet(
            "QHeaderView::section{font:11pt '微软雅黑'; color: white; font-weight: bold;}"
        )
        # 设置表格序数列字体粗体
        self.tableView.verticalHeader().setStyleSheet(
            "QHeaderView::section{font:11pt '微软雅黑'; color: white; font-weight: bold;}"
        )
        self.tableView.setSizeAdjustPolicy(QTableView.AdjustToContents)  # 设置 QTableView 自适应大小
        # 添加右键菜单
        self.tableView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableView.customContextMenuRequested.connect(self.open_menu)
        # 加载数据
        self.load_data()

    def load_data(self):
        model = QStandardItemModel(self)  # 创建一个 QStandardItemModel 作为数据模型
        # 设置模型的表头
        model.setHorizontalHeaderLabels(['变量名称', '备注'])
        variable_list = get_value_from_variable_table()  # 从数据库中获取数据
        # 添加数据到模型中
        for variable_tuple in variable_list:
            items = [QStandardItem(str(variable_tuple[0])), QStandardItem(str(variable_tuple[1]))]
            model.appendRow(items)
        # 将模型设置到 TableView 中
        self.tableView.setModel(model)
        self.tableView.resizeColumnToContents(1)
        self.tableView.setFocusPolicy(False)
        self.setCentralWidget(self.tableView)

    def open_menu(self, position):
        menu = QMenu()
        add_row_action = QAction("添加变量", self)
        delete_row_action = QAction("删除变量", self)
        menu.addAction(add_row_action)
        menu.addAction(delete_row_action)
        add_row_action.triggered.connect(self.add_row)
        delete_row_action.triggered.connect(self.delete_row)
        menu.exec_(self.tableView.viewport().mapToGlobal(position))

    def add_row(self):
        """添加一行到 TableView"""
        model = self.tableView.model()
        row_count = model.rowCount()
        model.insertRow(row_count)
        for column in range(model.columnCount()):
            model.setData(model.index(row_count, column), "New Data")

    def delete_row(self):
        """删除选定的行"""
        selection_model = self.tableView.selectionModel()
        rows = sorted(index.row() for index in selection_model.selectedIndexes())
        for row in reversed(rows):
            self.tableView.model().removeRow(row)

    def closeEvent(self, event):
        """关闭窗口时触发"""
        # 保存数据到数据库
        model = self.tableView.model()
        row_count = model.rowCount()
        variable_list = []
        for row in range(row_count):
            variable_name = model.index(row, 0).data()
            variable_remark = model.index(row, 1).data()
            variable_list.append((variable_name, variable_remark))
        # 保存数据到数据库
        set_value_to_variable_table(variable_list)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
