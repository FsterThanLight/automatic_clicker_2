from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QCursor
from PyQt5.QtWidgets import QDialog

from 变量池窗口 import VariablePool_Win
from 数据库操作 import extract_global_parameter, set_window_size, save_window_size, \
    get_variable_info
from 窗体.branchwin import Ui_branch


class Branch_exe_win(QDialog, Ui_branch):
    """弹出选择执行分支的窗体"""

    def __init__(self, parent=None, modes='分支选择'):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        self.modes = modes
        # 根据不同的模式设置窗体样式
        self.set_window_style(self.modes)

    def set_window_style(self, modes):
        """根据不同的模式设置窗体样式"""
        if modes == '分支选择':
            self.setWindowTitle('执行分支')
            self.label.setText('请选择要执行的分支')
            self.pushButton.setText('显示主窗口')
            # 绑定事件
            self.listView.doubleClicked.connect(self.open_select_option)  # 双击执行分支
            self.pushButton.clicked.connect(lambda: self.show_main(modes))
        elif modes == '变量选择':
            self.setWindowTitle('选择变量')
            self.label.setText('请选择要插入的变量')
            self.pushButton.setText('设置变量')
            # 绑定事件
            self.listView.doubleClicked.connect(self.write_to_textedit)
            self.pushButton.clicked.connect(lambda: self.show_main(modes))

    def load_lists(self, modes):
        """设置初始参数"""

        def add_listview(list_, listview):
            """添加listview"""
            model = QStandardItemModel()
            listview.setModel(model)
            for item in list_:
                model.appendRow(QStandardItem(item))

        if modes == '分支选择':
            branch_list = extract_global_parameter('分支表名')
            add_listview(branch_list, self.listView)
        elif modes == '变量选择':
            variable_list = get_variable_info('list')
            add_listview(variable_list, self.listView)

    def open_select_option(self):
        """打开选中的listview中的文件夹路径"""
        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text()  # 获取选中的值
            self.parent().comboBox.setCurrentText(value)  # 设置分支
            self.parent().start()  # 执行分支
            self.close()
        except Exception as e:
            print(e)

    def write_to_textedit(self):
        """将选中的值写入textedit，用于写入变量的模式"""
        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text()  # 获取选中的值
            self.parent().write_value_to_textedit(value)  # 设置分支
            self.close()
        except Exception as e:
            print(e)

    def show_main(self, modes='分支选择'):
        """显示主窗体"""
        if modes == '分支选择':
            # 如果父窗体最小化则显示
            if self.parent().isMinimized():
                self.parent().showNormal()
            self.close()
        elif modes == '变量选择':
            # 打开变量选择窗体
            variable_pool = VariablePool_Win(self)
            variable_pool.exec_()

    def showEvent(self, a0) -> None:
        # 设置窗口大小
        set_window_size(self)
        # 移动窗口到鼠标位置
        cursor_pos = QCursor.pos()
        # 移动窗口使窗口中心与鼠标位置重合
        self.move(cursor_pos.x() - self.width() / 2, cursor_pos.y() - self.height() / 2)
        self.load_lists(self.modes)  # 加载设置

    def closeEvent(self, event):
        # 保存窗体大小
        save_window_size((self.width(), self.height()), self.windowTitle())


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    window = Branch_exe_win()
    window.show()
    sys.exit(app.exec_())
