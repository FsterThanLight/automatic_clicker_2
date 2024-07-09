import time

from PyQt5.QtWidgets import QMessageBox

from 功能类 import get_available_path


def xxx_function(self, type_):
    """xxx的功能
    :param self:
    :param type_: 功能名称（按钮功能、主要功能）"""

    def get_parameters():
        """从tab页获取参数"""
        image_ = None
        parameter_1_ = None
        parameter_2_ = None
        parameter_3_ = None
        parameter_4_ = None
        # 检查参数是否有异常
        if image_ is None or parameter_1_ is None or parameter_2_ is None or parameter_3_ is None or parameter_4_ is None:
            QMessageBox.critical(self, "错误", "xxx")
            raise FileNotFoundError
        # 返回参数字典
        parameter_dic_ = {
            'x_1': parameter_1_,
            'x_2': parameter_2_,
            'x_3': parameter_3_,
            'x_4': parameter_4_,
        }
        return image_, parameter_dic_

    def put_parameters(image_, parameter_dic_):
        """将参数还原到tab页"""
        pass

    def test():
        """测试功能"""
        try:
            image_, parameter_dic_ = get_parameters()
            dic_ = self.get_test_dic(repeat_number_=int(self.spinBox.value()),
                                     image_=image_,
                                     parameter_1_=parameter_dic_
                                     )

            # 测试用例
            test_class = XxxxClss(self.out_mes, dic_)
            test_class.is_test = True
            test_class.start_execute()

        except Exception as e:
            print(e)
            self.out_mes.out_mes(f'指令错误请重试！', True)

    if type_ == '按钮功能':
        # 将不同的单选按钮添加到同一个按钮组
        # all_groupBoxes_ = [self.groupBox_22, self.groupBox_29]
        # for groupBox_ in all_groupBoxes_:
        #     groupBox_.clicked.connect(lambda _, gb=groupBox_: self.select_groupBox(gb, all_groupBoxes_))

        # self.lineEdit_3.setValidator(QIntValidator())  # 设置只能输入数字
        pass

    elif type_ == '写入参数':
        image, parameter_dic = get_parameters()
        # 将命令写入数据库
        func_info_dic = self.get_func_info()  # 获取功能区的参数
        self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                             repeat_number_=func_info_dic.get('重复次数'),
                                             exception_handling_=func_info_dic.get('异常处理'),
                                             image_=image,
                                             parameter_1_=parameter_dic,
                                             remarks_=func_info_dic.get('备注'))
    elif type_ == '加载信息':
        # 当t导航业显示时，加载信息到控件
        pass

    elif type_ == '还原参数':
        put_parameters(self.image_path, self.parameter_1)


class XxxxClss:
    """xxxx"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep: float = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic: dict = ins_dic  # 指令字典

        self.is_test: bool = False  # 是否测试
        self.cycle_number: int = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        re_try = self.ins_dic.get('重复次数')
        img = get_available_path(self.ins_dic.get('图像路径'), self.out_mes)
        parameter_dic_ = eval(self.ins_dic.get('参数1（键鼠指令）'))
        pass
        return re_try, None

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 解析指令字典
        reTry, x, = self.parsing_ins_dic()
        # 执行图像点击
        for _ in range(reTry):
            self.execute_func(x)
            if reTry > 1:
                time.sleep(self.time_sleep)

    def execute_func(self, x):
        pass
