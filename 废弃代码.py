# def check_update(self, warning):
#     """检查更新功能"""
#     # 获取下载地址、版本号、更新信息
#     list_1 = get_download_address(self, warning)
#     # print(list_1)
#     try:
#         address = list_1[0]
#         version = list_1[1]
#         information = list_1[2]
#         # 判断是否有更新
#         print(version)
#         if version != self.version:
#             x = QMessageBox.information(self, "更新检查",
#                                         "已发现最新版" + version + "\n是否手动下载最新安装包？" + '\n' + information,
#                                         QMessageBox.Yes | QMessageBox.No,
#                                         QMessageBox.Yes)
#             if x == QMessageBox.Yes:
#                 # 打开下载地址
#                 webbrowser.open(address)
#                 # os.popen('update.exe')
#                 sys.exit()
#         else:
#             if warning == 1:
#                 QMessageBox.information(self, "更新检查", "当前" + self.version + "已是最新版本。")
#             else:
#                 pass
#     except TypeError:
#         pass


# headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
#                          'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36'}
#
#
# def load_json():
#     """从json文件中加载更新网址和保留文件名"""
#     file_name = 'update_data_dic.json'
#     with open(file_name, 'r', encoding='utf8') as f:
#         data = json.load(f)
#     url = cryptocode.decrypt(data['url_encrypt'], '123456')
#     return url
#
#
# def get_download_address(main_window_, warning):
#     """获取下载地址、版本信息、更新说明"""
#     global headers
#     url = load_json()
#     try:
#         res = requests.get(url, headers=headers, timeout=0.2)
#         info = cryptocode.decrypt(res.text, '123456')
#         list_1 = info.split('=')
#         return list_1
#     except requests.exceptions.ConnectionError:
#         if warning == 1:
#             # print("无法获取更新信息，请检查网络。")
#             QMessageBox.critical(main_window_, "更新检查", "无法获取更新信息，请检查网络。")
#             time.sleep(1)
#         else:
#             pass

# class Login(QWidget, Ui_Login):
#     """登录窗体"""
#
#     def __init__(self):
#         super(Login, self).__init__()
#         self.setupUi(self)
#         # 登录按钮
#         self.pushButton.clicked.connect(self.login_main_window)
#         self.lineEdit_2.returnPressed.connect(self.login_main_window)
#         self.lineEdit.returnPressed.connect(self.lineEdit_2.setFocus)
#
#     def login_main_window(self):
#         """登录进主窗口"""
#         # 连接数据库
#         cursor, con = sqlitedb()
#         # 获取数据库中的用户名和密码
#         cursor.execute('select 账号,密码 from 账户')
#         list_account = cursor.fetchall()
#         close_database(cursor, con)
#         # 判断登录
#         ac = (self.lineEdit.text(), self.lineEdit_2.text())
#         if ac in list_account:
#             self.close()
#             # 如果选中记住密码则保存账户id
#             if self.checkBox.isChecked():
#                 cursor, con = sqlitedb()
#                 # 根据账号和密码获取id
#                 cursor.execute('select ID from 账户 where 账号=? and 密码=?', (ac[0], ac[1]))
#                 account_id = cursor.fetchall()[0][0]
#                 cursor.execute('update 设置 set 值 = ? where 设置类型=?', (str(account_id), '账户ID'))
#                 cursor.execute('update 设置 set 值 = ? where 设置类型=?', (1, '记住密码'))
#                 con.commit()
#                 close_database(cursor, con)
#             elif not self.checkBox.isChecked():
#                 cursor, con = sqlitedb()
#                 cursor.execute('update 设置 set 值 = ? where 设置类型=?', (0, '记住密码'))
#                 con.commit()
#                 close_database(cursor, con)
#             # 创建主窗体
#             main_window_ = Main_window()
#             # # 显示窗体，并根据设置检查更新
#             main_window_.main_show()
#         else:
#             QMessageBox.information(self, '提示', '密码错误。')
#
#     def login_show(self):
#         """显示登录窗体"""
#         cursor, con = sqlitedb()
#         cursor.execute('select 值 from 设置 where 设置类型=?', ('记住密码',))
#         remember_password = cursor.fetchall()[0][0]
#         cursor.execute('select 值 from 设置 where 设置类型=?', ('账户ID',))
#         account_id = cursor.fetchall()[0][0]
#         close_database(cursor, con)
#         self.show()
#         if remember_password == 1:
#             self.checkBox.setChecked(True)
#             cursor, con = sqlitedb()
#             cursor.execute('select 账号,密码 from 账户 where ID=?', (account_id,))
#             account = cursor.fetchall()[0]
#             close_database(cursor, con)
#             self.lineEdit.setText(account[0])
#             self.lineEdit_2.setText(account[1])
#             self.lineEdit_2.setFocus()
#         else:
#             self.lineEdit.setFocus()
