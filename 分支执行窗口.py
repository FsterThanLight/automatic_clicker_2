import ctypes

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QDialog, QHeaderView

from ini控制 import get_branch_info, set_window_size, save_window_size
from 窗体.分支执行 import Ui_Branch


class BranchWindow(QDialog, Ui_Branch):
    """分支执行窗口"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.doubleClicked.connect(self.open_select_option)
        self.tableWidget.installEventFilter(self)  # 安装事件过滤器,重新设置表格的快捷键

    def load_branch_data(self):
        """加载分支数据"""
        branch_info = get_branch_info()
        self.tableWidget.setRowCount(len(branch_info))
        for row, (name, short_desc, repeat_times) in enumerate(branch_info):
            for col, text in enumerate((name, short_desc)):
                item = QtWidgets.QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                self.tableWidget.setItem(row, col, item)
            # 在第三列添加QSpinBox控件
            spin_box = QtWidgets.QSpinBox()
            # 设置最大值
            spin_box.setMaximum(1000000)
            spin_box.setMinimum(-1)
            spin_box.setValue(repeat_times)
            spin_box.setAlignment(Qt.AlignCenter)
            self.tableWidget.setCellWidget(row, 2, spin_box)
        # 将焦点设置到第一行
        self.tableWidget.setCurrentCell(0, 0)

    @staticmethod
    def set_caps_lock_status(judge: str):
        """设置大写锁定状态，防止快捷键失效
        :param judge: open or close
        """
        try:
            VK_CAPITAL = 0x14
            KEYEVENTF_EXTENDEDKEY = 0x0001
            KEYEVENTF_KEYUP = 0x0002
            # 获取大写锁定状态
            hllDll = ctypes.WinDLL("User32.dll")
            caps_lock_state = hllDll.GetKeyState(VK_CAPITAL)
            is_caps_lock_on = caps_lock_state & 1
            if judge == 'close' and is_caps_lock_on:
                # 如果大写锁定打开，则关闭
                ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0)
                ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
            elif judge == 'open' and not is_caps_lock_on:
                # 如果大写锁定关闭，则打开
                ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0)
                ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
        except Exception as e:
            print('设置大写锁定状态错误！', e)

    def open_select_option(self):
        """打开选中的分支"""
        selected_row = self.tableWidget.currentRow()
        try:
            if selected_row != -1:
                branch_name = self.tableWidget.item(selected_row, 0).text()
                repeat_times = self.tableWidget.cellWidget(selected_row, 2).value()
                print(f'执行分支: {branch_name}，重复次数: {repeat_times}')
                self.parent().start_from_branch(branch_name, repeat_times)
                self.close()
        except Exception as e:
            print(e)

    def trigger_using_number_keys(self, number):
        """设置到对应的行"""
        if number <= self.tableWidget.rowCount():
            self.tableWidget.setCurrentCell(number - 1, 0)
            self.open_select_option()

    def showEvent(self, a0) -> None:
        # 设置窗口大小
        set_window_size(self)
        # 移动窗口到鼠标位置
        cursor_pos = QCursor.pos()
        # 移动窗口使窗口中心与鼠标位置重合
        self.move(cursor_pos.x() - int(self.width() / 2), cursor_pos.y() - int(self.height() / 2))
        # 加载分支数据
        self.load_branch_data()
        self.set_caps_lock_status('open')  # 打开大写锁定

    def closeEvent(self, event):
        # 保存窗体大小
        save_window_size(self.width(), self.height(), self.windowTitle())
        self.set_caps_lock_status('close')  # 关闭大写锁定

    def eventFilter(self, obj, event):
        # 重写self.tableWidget的快捷键事件
        if obj == self.tableWidget:
            if event.type() == 6:  # 键盘按下事件
                branch_info = get_branch_info()
                for i, (name, key_str, repeat_times) in enumerate(branch_info):
                    if event.key() == self.key_name_to_qt_key(key_str):
                        self.trigger_using_number_keys(i + 1)
        return super().eventFilter(obj, event)

    @staticmethod
    def key_name_to_qt_key(key_name) -> int:
        """Convert a key name to its corresponding Qt key value."""
        key_mapping = {
            '0': Qt.Key_0,
            '1': Qt.Key_1,
            '2': Qt.Key_2,
            '3': Qt.Key_3,
            '4': Qt.Key_4,
            '5': Qt.Key_5,
            '6': Qt.Key_6,
            '7': Qt.Key_7,
            '8': Qt.Key_8,
            '9': Qt.Key_9,
            'A': Qt.Key_A,
            'B': Qt.Key_B,
            'C': Qt.Key_C,
            'D': Qt.Key_D,
            'E': Qt.Key_E,
            'F': Qt.Key_F,
            'G': Qt.Key_G,
            'H': Qt.Key_H,
            'I': Qt.Key_I,
            'J': Qt.Key_J,
            'K': Qt.Key_K,
            'L': Qt.Key_L,
            'M': Qt.Key_M,
            'N': Qt.Key_N,
            'O': Qt.Key_O,
            'P': Qt.Key_P,
            'Q': Qt.Key_Q,
            'R': Qt.Key_R,
            'S': Qt.Key_S,
            'T': Qt.Key_T,
            'U': Qt.Key_U,
            'V': Qt.Key_V,
            'W': Qt.Key_W,
            'X': Qt.Key_X,
            'Y': Qt.Key_Y,
            'Z': Qt.Key_Z,
            'F1': Qt.Key_F1,
            'F2': Qt.Key_F2,
            'F3': Qt.Key_F3,
            'F4': Qt.Key_F4,
            'F5': Qt.Key_F5,
            'F6': Qt.Key_F6,
            'F7': Qt.Key_F7,
            'F8': Qt.Key_F8,
            'F9': Qt.Key_F9,
            'F10': Qt.Key_F10,
            'F11': Qt.Key_F11,
            'F12': Qt.Key_F12,
            'Esc': Qt.Key_Escape,
            'Return': Qt.Key_Return,
            'Space': Qt.Key_Space,
            'Tab': Qt.Key_Tab,
            'Backspace': Qt.Key_Backspace,
            'Ins': Qt.Key_Insert,
            'Del': Qt.Key_Delete,
            'Home': Qt.Key_Home,
            'End': Qt.Key_End,
            'PgUp': Qt.Key_PageUp,
            'PgDown': Qt.Key_PageDown,
            'Left': Qt.Key_Left,
            'Up': Qt.Key_Up,
            'Right': Qt.Key_Right,
            'Down': Qt.Key_Down,
            'Shift': Qt.Key_Shift,
            'Control': Qt.Key_Control,
            'Alt': Qt.Key_Alt,
            'Meta': Qt.Key_Meta,
            '.': Qt.Key_Period,
            ',': Qt.Key_Comma,
            '/': Qt.Key_Slash,
            '\\': Qt.Key_Backslash,
            '[': Qt.Key_BracketLeft,
            ']': Qt.Key_BracketRight,
            '-': Qt.Key_Minus,
            '=': Qt.Key_Equal,
            '`': Qt.Key_QuoteLeft,
            ';': Qt.Key_Semicolon,
            "'": Qt.Key_Apostrophe,
            '{': Qt.Key_BraceLeft,
            '}': Qt.Key_BraceRight,
            'Ctrl': Qt.Key_Control,
        }
        return key_mapping.get(key_name, None)


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    window = BranchWindow()
    window.show()
    sys.exit(app.exec_())
