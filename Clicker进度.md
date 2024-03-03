- [ ] 支付系统
- [ ] 软件更新系统
- [x] 加密软件
- [ ] 服务器和网站
- [x] UI美化
- [ ] 功能完整
- [ ] 编译为python功能

```text
nuitka --standalone --windows-disable-console --enable-plugin=pyqt5 --plugin-enable=tk-inter --windows-icon-from-ico=clicker.ico --include-package=pygments --include-package=pyttsx4 --include-package-data=selenium --remove-output Clicker.py
```

```text
nuitka --standalone --onefile --windows-disable-console --enable-plugin=pyqt5 --plugin-enable=tk-inter --windows-icon-from-ico=clicker.ico --include-package=pygments --include-package=pyttsx4 --include-package-data=selenium --remove-output Clicker.py
```
