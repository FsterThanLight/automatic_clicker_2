import time


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
        return image_, parameter_1_, parameter_2_, parameter_3_, parameter_4_

    def test():
        """测试功能"""
        try:
            image_, parameter_1_, parameter_2_, parameter_3_, parameter_4_ = get_parameters()
            dic_ = self.get_test_dic(
                image_,
                parameter_1_,
                parameter_2_,
                parameter_3_,
                parameter_4_,
            )

            # 测试用例
            test_class = XxxxClss(self.out_mes, dic_)
            test_class.is_test = True
            test_class.start_execute()

        except Exception as e:
            print(e)
            self.out_mes.out_mes(f'xxxx！', True)

    if type_ == '按钮功能':
        pass

    elif type_ == '写入参数':
        image, parameter_1, parameter_2, parameter_3, parameter_4 = get_parameters()
        # 将命令写入数据库
        func_info_dic = self.get_func_info()  # 获取功能区的参数
        self.writes_commands_to_the_database(instruction_=func_info_dic.get('指令类型'),
                                             repeat_number_=func_info_dic.get('重复次数'),
                                             exception_handling_=func_info_dic.get('异常处理'),
                                             image_=image,
                                             parameter_1_=parameter_1,
                                             parameter_2_=parameter_2,
                                             parameter_3_=parameter_3,
                                             parameter_4_=parameter_4,
                                             remarks_=func_info_dic.get('备注'))
    elif type_ == '加载信息':
        # 当t导航业显示时，加载信息到控件
        pass


class XxxxClss:
    """图像点击"""

    def __init__(self, outputmessage, ins_dic, cycle_number=1):
        # 设置参数
        self.time_sleep = 0.5  # 等待时间
        self.out_mes = outputmessage  # 用于输出信息到不同的窗口
        self.ins_dic = ins_dic  # 指令字典

        self.is_test = False  # 是否测试
        self.cycle_number = cycle_number  # 循环次数

    def parsing_ins_dic(self):
        """从指令字典中解析出指令参数"""
        re_try = self.ins_dic.get('重复次数')
        pass
        return re_try, None

    def start_execute(self):
        """开始执行鼠标点击事件"""
        # 解析指令字典
        reTry, x, = self.parsing_ins_dic()
        # 执行图像点击
        if reTry == 1:
            self.execute_func(x)
        elif reTry > 1:
            i = 1
            while i < reTry + 1:
                self.execute_func(x)
                i += 1
                time.sleep(self.time_sleep)

    def execute_func(self, x):
        pass
