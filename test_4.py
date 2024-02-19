class WaitWindow(QDialog, Ui_wait_win):
    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        super().__init__()
        self.setupUi(self)
        # 倒计时
        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update_label)  # 每次计时结束，触发update_label
        # 窗口参数
        self.out_mes = outputmessage  # 用于输出信息
        self.ins_dic = ins_dic  # 指令字典
        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数
        # 立即执行
        self.is_raise = False
        self.pushButton.clicked.connect(self.stop_win)  # 停止运行
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.WindowCloseButtonHint)  # 置顶窗口, 禁用窗口最大化和最小化

    def update_label(self):
        """更新倒计时"""
        current_count = int(self.label_2.text())
        current_count -= 1
        if current_count < 1:
            self.timer.stop()
            self.close()
            self.out_mes.out_mes('已结束等待窗口', self.is_test)
            return
        self.label_2.setText(str(current_count))
        self.out_mes.out_mes('倒计时：%s' % current_count, self.is_test)

    def parsing_ins_dic(self):
        """解析指令字典"""
        list_dic = {
            '窗口标题': self.ins_dic.get('参数1（键鼠指令）'),
            '提示信息': self.ins_dic.get('参数2'),
            '等待时间': self.ins_dic.get('参数3')
        }
        return list_dic

    def stop_win(self):
        """停止窗口"""
        self.timer.stop()
        self.close()
        self.out_mes.out_mes('已结束等待窗口', self.is_test)

    def start_execute(self):
        """显示窗口"""
        list_dic = self.parsing_ins_dic()
        self.label_2.setText(str(list_dic.get('等待时间')))  # 重置倒计时
        self.setWindowTitle(list_dic.get('窗口标题'))  # 设置窗口标题
        self.label.setText(list_dic.get('提示信息'))  # 设置提示信息
        self.out_mes.out_mes('弹出等待窗口', self.is_test)
        # 开始倒计时
        self.timer.start()
        self.exec_()