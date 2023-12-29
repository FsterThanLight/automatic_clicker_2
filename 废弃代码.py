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
#     file_name = 'update_data.json'
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
