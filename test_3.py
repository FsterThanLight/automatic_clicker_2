from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QDialog, QVBoxLayout


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Main Window")

        self.central_widget = QPushButton("Open Sub Window", self)
        self.central_widget.clicked.connect(self.open_sub_window)
        self.setCentralWidget(self.central_widget)

    def open_sub_window(self):
        sub_window = SubWindow(self)
        sub_window.setWindowTitle("Sub Window")
        layout = QVBoxLayout(sub_window)
        layout.addWidget(QPushButton("Button in Sub Window"))
        # 将窗口设置为模态，以防止同时操作主窗口和子窗口
        sub_window.setModal(True)
        sub_window.exec_()


class SubWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)


if __name__ == '__main__':
    app = QApplication([])
    main_window = MainWindow()
    main_window.show()
    app.exec_()
