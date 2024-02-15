import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("菜单栏悬停示例")

        # 创建顶层菜单
        menubar = self.menuBar()

        # 创建一级菜单
        file_menu = menubar.addMenu('文件')

        # 创建二级菜单
        new_menu = QMenu('新建', self)
        new_menu.addAction('文件')
        new_menu.addAction('文件夹')

        # 添加二级菜单到一级菜单
        file_menu.addMenu(new_menu)

        # 创建动作
        exit_action = QAction('退出', self)
        exit_action.triggered.connect(self.close)

        # 将动作添加到一级菜单
        file_menu.addAction(exit_action)

        # 监听鼠标悬停事件
        exit_action.hovered.connect(self.onHovered)

        # 显示窗口
        self.show()

    def onHovered(self):
        # 获取信号源
        action = self.sender()
        if action.menu():
            # 如果该action有菜单，则打开菜单
            action.menu().exec_(action.geometry().bottomLeft())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())
