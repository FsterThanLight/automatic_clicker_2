import pymsgbox


# pymsgbox.alert(text='', title='', button='OK')
# pymsgbox.confirm(text='', title='', buttons=['OK', 'Cancel'])
# pymsgbox.prompt(text='', title='', default='')
# pymsgbox.password(text='', title='', default='', mask='*')
#
# # 弹窗确认框
# pymsgbox.alert('This is an alert.', 'Alert!')
#
# # 选择确认框
# pymsgbox.confirm('你是否要查看以下内容?', '查看确认', ["确定", '取消'])
#
# # 密码输入框,mask指定密码代替符号
# res = pymsgbox.password('Enter your password.', mask='$')
# print(res)
#
# # 默认输入框
# pymsgbox.prompt('What does the fox say?', default='This reference dates this example.')
#
# # 选择确认框，设置时间后自动消失
# pymsgbox.confirm('你是否要查看以下内容?', '查看确认', ["确定", '取消'], timeout=2000)

def alert_dialog_box(text, title, icon_):
    """测试功能
    :param text: 弹窗内容
    :param title: 弹窗标题
    :param icon_: 弹窗图标"""
    icon_dic = {
        'STOP': pymsgbox.STOP,
        'WARNING': pymsgbox.WARNING,
        'INFO': pymsgbox.INFO,
        'QUESTION': pymsgbox.QUESTION,
    }
    pymsgbox.alert(
        text=text,
        title=title,
        icon=icon_dic.get(icon_)
    )


def confirm_dialog_box(text, title, buttons):
    """测试功能
    :param text: 弹窗内容
    :param title: 弹窗标题
    :param buttons: 按钮"""
    pymsgbox.confirm(
        text=text,
        title=title,
        buttons=buttons
    )


if __name__ == '__main__':
    # alert_dialog_box('This is an alert.', 'Alert!', 'STOP')
    confirm_dialog_box('你是否要查看以下内容?', '查看确认', ["确定", '取消'])