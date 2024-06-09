from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QDialog
from 窗体.分支执行 import Ui_Branch


class BranchWindow(QDialog, Ui_Branch):
    """分支执行窗口"""

    def __init__(self):
        super().__init__(parent=None)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        # 加载分支数据
        self.load_branch_data()

    def load_branch_data(self):
        """加载分支数据"""
        # 在表格中写入数据
        branch_list = ['分支1', '分支2', '分支3', '分支4', '分支5']
        shortcut_key = ['Ctrl+1', 'Ctrl+2', 'Ctrl+3', 'Ctrl+4', 'Ctrl+5']
        self.tableWidget.setRowCount(len(branch_list))
        for i in range(len(branch_list)):
            self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(branch_list[i]))
            # 使用 QKeySequenceEdit 控件显示快捷键
            key_sequence = QKeySequence(shortcut_key[i])
            key_sequence_edit = QtWidgets.QKeySequenceEdit(key_sequence)
            self.tableWidget.setCellWidget(i, 1, key_sequence_edit)


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    window = BranchWindow()
    window.show()
    sys.exit(app.exec_())