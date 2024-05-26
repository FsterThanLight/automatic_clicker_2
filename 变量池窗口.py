import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QDialog, QMenu, QAction, QStyle, QApplication

from 数据库操作 import get_value_from_variable_table, set_value_to_variable_table
from 窗体.variablepool import Ui_VariablePool


class VariablePool_Win(QDialog, Ui_VariablePool):
    """变量池窗体"""

    def __init__(self, parent=None):
        super().__init__(parent)
        # 初始化变量池窗口
        self.setupUi(self)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )  # 隐藏帮助按钮
        self.set_style()  # 设置窗体样式
        # 添加右键菜单
        self.tableView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableView.customContextMenuRequested.connect(self.open_menu)
        self.load_data()  # 加载数据

    def set_style(self):
        """设置窗体样式"""
        # 设置表格风格
        self.tableView.horizontalHeader().setStretchLastSection(
            True
        )  # 设置最后一列拉伸至最大
        self.tableView.setStyleSheet(
            "QHeaderView::section{background:red;}"
        )  # 设置表头背景色
        # 设置标题字体粗体
        self.tableView.horizontalHeader().setStyleSheet(
            "QHeaderView::section{font:11pt '微软雅黑'; color: white; font-weight: bold;}"
        )
        # 设置表格序数列字体粗体
        self.tableView.verticalHeader().setStyleSheet(
            "QHeaderView::section{font:11pt '微软雅黑'; color: white; font-weight: bold;}"
        )

    def load_data(self):
        model = QStandardItemModel(self)  # 创建一个 QStandardItemModel 作为数据模型
        # 设置模型的表头
        model.setHorizontalHeaderLabels(["变量名称", "备注", "值"])
        variable_list = get_value_from_variable_table()  # 从数据库中获取数据
        # 添加数据到模型中
        for variable_tuple in variable_list:
            items = [
                QStandardItem(str(variable_tuple[0])),
                QStandardItem(str(variable_tuple[1])),
                QStandardItem(str(variable_tuple[2])),
            ]
            model.appendRow(items)
        # 将模型设置到 TableView 中
        self.tableView.setModel(model)
        self.tableView.resizeColumnToContents(1)
        self.tableView.setFocusPolicy(False)
        # 重新设置窗口大小以适应表格的大小
        width = self.tableView.horizontalHeader().length() + 100  # 加上一些额外空间
        height = self.tableView.verticalHeader().length() + 100
        self.resize(width, height)  # 设置窗口大小为表格大小加上一些额外空间

    def open_menu(self, position):
        menu = QMenu()
        add_row_action = QAction("添加变量", self)
        delete_row_action = QAction("删除变量", self)
        menu.addAction(add_row_action)
        menu.addAction(delete_row_action)
        # 设置图标
        add_row_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        delete_row_action.setIcon(
            self.style().standardIcon(QStyle.SP_DialogDiscardButton)
        )
        # 绑定事件
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
            variable_value = model.index(row, 2).data()
            variable_list.append((variable_name, variable_remark, variable_value))
        # 保存数据到数据库
        # print(variable_list)
        set_value_to_variable_table(variable_list)

        # 父窗口加载数据
        if self.parent():
            try:  # 重新加载父窗口的数据，用于选择窗口的变量更新
                self.parent().load_lists("变量选择")
            except AttributeError:
                pass
            try:  # 重新加载父窗口的数据，用于导航窗口的变量更新
                self.parent().tab_widget_change()
            except AttributeError:
                pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = VariablePool_Win()
    win.show()
    sys.exit(app.exec_())
