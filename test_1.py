import sys

from PyQt5.QtWidgets import QApplication


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


from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QDialog, QHeaderView, QTableWidgetItem

class ShortcutTable(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        # Set window title
        self.setWindowTitle("快捷键说明")
        self.table = QtWidgets.QTableWidget()
        self.table.setRowCount(12)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["快捷键", "说明"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.setWindowFlags(
            self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint
        )  # 隐藏帮助按钮
        self.button = QtWidgets.QPushButton("我知道了")
        self.button.clicked.connect(self.close)
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
        for row, (shortcut, description) in enumerate(data):
            shortcut_item = QTableWidgetItem(shortcut)
            shortcut_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
            shortcut_item.setTextAlignment(QtCore.Qt.AlignCenter)
            description_item = QTableWidgetItem(description)
            description_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
            description_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table.setItem(row, 0, shortcut_item)
            self.table.setItem(row, 1, description_item)

        # Set layout
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.button)
        self.setLayout(layout)
        # 获取表格的总高度，设置窗口的高度
        table_height = self.table.verticalHeader().length()
        self.resize(400, table_height + 150)
        self.table.setFocusPolicy(QtCore.Qt.NoFocus)

        # Apply QSS
        self.setStyleSheet("""
            QTableWidget {
                border: none;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: transparent;
            }
            QTableWidget::item:focus {
                background-color: transparent;
            }
            QTableWidget::item:nth-child(1) {
                background-color: #E1ECF4;
                color: #1D6FAD;
                border: 1px solid #B7D7F0;
                border-radius: 10px;
                padding: 5px;
                margin: 2px;
                text-align: center;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px 24px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = ShortcutTable()
    window.show()
    sys.exit(app.exec_())
