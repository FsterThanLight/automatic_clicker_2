import pymsgbox
import winsound

# xx = pymsgbox.alert(text='This is an alert!',
#                     title='Alert',
#                     icon=2)
# print(xx)

# xx = pymsgbox.confirm(text='Do you want to continue?',
#                       title='Confirm',
#                       # buttons=[pymsgbox.YES_TEXT, pymsgbox.NO_TEXT, pymsgbox.CANCEL_TEXT])
#                       buttons=[pymsgbox.ABORT_TEXT, pymsgbox.RETRY_TEXT, pymsgbox.IGNORE_TEXT])
# print(xx)
for i in range(3):
    winsound.Beep(500, 300)
