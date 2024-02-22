import sys
from PyQt5.QtWidgets import QTextEdit, QApplication


def substitution_variable(text: str, value_dict: dict) -> str:
    """将text中的@@variable@@替换为value"""
    for key, value in value_dict.items():
        text = text.replace(f'@@{key}@@', value)
    return text


def append_textedit(textEdit_, new_text):
    # Create formatted strings
    errorFormat_ = '<font color="red">{}</font>'
    # 使textEdit显示不同的文本
    textEdit_.insertHtml((errorFormat_.format(new_text)))


if __name__ == '__main__':
    app = QApplication([])

    textEdit = QTextEdit()

    # Create formatted strings
    errorFormat = '<font color="red" size="50">{}</font>'
    # warningFormat = '<font color="orange" size="50">{}</font>'
    # validFormat = '<font color="green" size="50">{}</font>'

    # textEdit.append('sjalijglisajldk')
    # # 插入富文本
    # textEdit.insertHtml(errorFormat.format_('Error'))
    # # 插入普通文本
    # textEdit.insertPlainText('sjalijglisajldk')
    textEdit.insertHtml('普通文本1')
    append_textedit(textEdit, 'Error')
    textEdit.insertHtml('普通文本2')

    textEdit.show()
    print(textEdit.toPlainText())
    sys.exit(app.exec_())
