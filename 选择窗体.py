from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QCursor
from PyQt5.QtWidgets import QDialog, QHeaderView, QTableWidgetItem

from ini控制 import set_window_size, save_window_size, get_branch_info
from 变量池窗口 import VariablePool_Win
from 数据库操作 import get_variable_info
from 窗体.branchwin import Ui_branch


class Variable_selection_win(QDialog, Ui_branch):
    """弹出选择执行分支的窗体"""

    def __init__(self, parent=None, modes='分支选择'):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        self.modes = modes
        # 根据不同的模式设置窗体样式
        self.set_window_style(self.modes)
        self.listView.installEventFilter(self)  # 安装事件过滤器,重新设置表格的快捷键

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
                # item前面加上数字序号
                item = str(list_.index(item) + 1) + '. ' + item
                model.appendRow(QStandardItem(item))

        if modes == '分支选择':
            branch_list = get_branch_info(True)
            add_listview(branch_list, self.listView)
        elif modes == '变量选择':
            variable_list = get_variable_info('list')
            add_listview(variable_list, self.listView)

    def open_select_option(self):
        """打开选中的listview中的文件夹路径"""
        try:
            indexes = self.listView.selectedIndexes()
            if indexes:
                selected_text = indexes[0].data().split('. ')[1]  # 直接获取选中项的文本值
                self.parent().comboBox.setCurrentText(selected_text)  # 设置分支
                self.parent().start()  # 执行分支
                self.close()
        except Exception as e:
            print(e)

    def trigger_using_number_keys(self, number):
        """设置到对应的行"""
        if number <= self.listView.model().rowCount():
            self.listView.setCurrentIndex(self.listView.model().index(number - 1, 0))
            self.open_select_option()  # 触发双击事件

    def write_to_textedit(self):
        """将选中的值写入textedit，用于写入变量的模式"""
        try:
            indexes = self.listView.selectedIndexes()
            value = self.listView.model().itemFromIndex(indexes[0]).text().split('. ')[1]  # 获取选中的值
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
        self.move(cursor_pos.x() - int(self.width() / 2), cursor_pos.y() - int(self.height() / 2))
        self.load_lists(self.modes)  # 加载设置

    def closeEvent(self, event):
        # 保存窗体大小
        save_window_size(self.width(), self.height(), self.windowTitle())

    def eventFilter(self, obj, event):
        # 重写self.tableWidget的快捷键事件
        if obj == self.listView:
            if event.type() == 6:  # 键盘按下事件
                key_to_row = {  # 数字键对应的行
                    Qt.Key_1: 1,
                    Qt.Key_2: 2,
                    Qt.Key_3: 3,
                    Qt.Key_4: 4,
                    Qt.Key_5: 5,
                    Qt.Key_6: 6,
                    Qt.Key_7: 7,
                    Qt.Key_8: 8,
                    Qt.Key_9: 9,
                }
                # 检查事件的键是否在字典中
                if event.key() in key_to_row:
                    row_number = key_to_row[event.key()]
                    self.trigger_using_number_keys(row_number)  # 使用数字键触发对应的行
        return super().eventFilter(obj, event)


class ShortcutTable(QDialog):
    def __init__(self, parent=None, title=None, data=None, width=300):
        super().__init__(parent)

        # Set window title
        self.setWindowTitle("快捷键说明")
        self.table = QtWidgets.QTableWidget()
        self.table.setRowCount(12)  # 调整行数，确保足够显示所有数据
        self.table.setColumnCount(2)
        if title:
            self.table.setHorizontalHeaderLabels(title)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )  # 隐藏帮助按钮

        self.button = QtWidgets.QPushButton("我知道了")
        self.button.clicked.connect(self.close)

        if data:
            self.table.setRowCount(len(data))
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
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.button)
        self.setLayout(layout)

        # 获取表格的总高度，设置窗口的高度
        table_height = self.table.verticalHeader().length()
        self.resize(width, table_height + 150)
        self.table.setFocusPolicy(Qt.NoFocus)


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    window = Variable_selection_win()
    window.show()
    sys.exit(app.exec_())
