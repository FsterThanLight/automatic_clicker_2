import sys

from PyQt5 import QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QVBoxLayout, QApplication, QWidget, QHeaderView


def timer(func):
    def func_wrapper(*args, **kwargs):
        from time import time
        time_start = time()
        result = func(*args, **kwargs)
        time_end = time()
        time_spend = time_end - time_start
        print('%s cost time: %.3f s' % (func.__name__, time_spend))
        return result

    return func_wrapper


class ShortcutTable(QWidget):
    def __init__(self):
        super().__init__()

        # Set window title
        self.setWindowTitle("快捷键说明")

        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/按钮图标/窗体/res/图标.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)

        # Set up the table
        self.table = QTableWidget()
        self.table.setRowCount(12)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["快捷键", "说明"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        # 禁用最大化按钮、禁用最小化按钮
        self.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowStaysOnTopHint)

        # Data to be displayed in the table
        data = [
            ("Ctrl+Enter", "添加指令"),
            ("Ctrl+C", "复制指令"),
            ("Delete", "删除指令"),
            ("Shift+↑", "上移指令"),
            ("Shift+↓", "下移指令"),
            ("Ctrl+↑", "切换到上个分支"),
            ("Ctrl+↓", "切换到下个分支"),
            ("Ctrl+G", "转到分支"),
            ("Ctrl+Y", "修改指令"),
            ("Ctrl+D", "导入指令"),
            ("Ctrl+S", "保存指令"),
            ("Ctrl+Alt+S", "另存为Excel")
        ]

        # Populate the table with data
        for row, (shortcut, description) in enumerate(data):
            shortcut_item = QTableWidgetItem(shortcut)
            shortcut_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            shortcut_item.setTextAlignment(Qt.AlignCenter)
            description_item = QTableWidgetItem(description)
            description_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            description_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, shortcut_item)
            self.table.setItem(row, 1, description_item)

        # Set layout
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        self.setLayout(layout)
        # 获取表格的总高度，设置窗口的高度
        table_height = self.table.verticalHeader().length()
        self.resize(300, table_height + 100)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ShortcutTable()
    window.show()
    sys.exit(app.exec_())
