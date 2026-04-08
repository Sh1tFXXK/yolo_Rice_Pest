# -*- coding: utf-8 -*-
import sys
import random
import string
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from db_helper import check_user
from RegisterDialog import RegisterDialog
from config_text import LOGIN_LEFT_TEXTS
from QMessageBox_helper import show_QMessageBox

# 导入自动生成的界面类
from login_layout import Ui_LoginWindow

class LoginMainWindow(QtWidgets.QWidget, Ui_LoginWindow):
    def __init__(self, user_info=None):
        super().__init__()
        # 初始化UI界面
        self.setupUi(self)
        
        # 先生成验证码
        self.captcha_code = self.generate_captcha()
        self.init_ui()

    def init_ui(self):
        """初始化UI界面。"""
        # 注意：窗口标题、大小和样式已经在UI文件中设置，这里不需要重复设置
        self.setWindowTitle(LOGIN_LEFT_TEXTS["title"] + " - 登录")
        # self.setFixedSize(1000, 600)
        # self.setStyleSheet("background: white;")

        self.title.setText(LOGIN_LEFT_TEXTS["title"])
        self.login_btn.setText("进入" + LOGIN_LEFT_TEXTS["title"])
        logo_pixmap = QPixmap("icon/logo.png")
        scaled_pixmap = logo_pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.logo_label.setPixmap(scaled_pixmap)
        self.subtitle.setText(LOGIN_LEFT_TEXTS["subtitle"])
        self.description.setText(LOGIN_LEFT_TEXTS["description"])
        self.login_title.setText(LOGIN_LEFT_TEXTS["title"] + "登录")
        
        # 直接使用已生成的验证码
        self.captcha_display.setText(self.captcha_code)

    def refresh_captcha(self):
        """刷新验证码。"""
        self.captcha_code = self.generate_captcha()
        self.captcha_display.setText(self.captcha_code)
        self.captcha_input.clear()

    def generate_captcha(self):
        """生成一个4位的随机字母和数字组合的验证码。"""
        self.captcha_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
        return self.captcha_code  # 添加返回语句

    def show_register_dialog(self):
        """显示用户注册对话框。"""
        dialog = RegisterDialog(self)
        dialog.exec_()

    def login(self):
        """
        处理登录按钮的点击事件。
        """
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        captcha = self.captcha_input.text().strip().upper()
        
        if not username:
            show_QMessageBox("登录失败", "请输入用户名！", QMessageBox.Warning)
            return
            
        if not password:
            show_QMessageBox("登录失败", "请输入密码！", QMessageBox.Warning)
            return
            
        if not captcha:
            show_QMessageBox("登录失败", "请输入验证码！", QMessageBox.Warning)
            return
            
        if captcha != self.captcha_code:
            show_QMessageBox("登录失败", "验证码错误！", QMessageBox.Warning)
            self.refresh_captcha()
            return
        
        # 校验数据库用户名和密码
        user_info = check_user(username, password)
        if user_info:
            self.show_main_system(user_info)
        else:
            show_QMessageBox("登录失败", "用户名或密码错误！", QMessageBox.Warning)
        self.refresh_captcha()

    def show_main_system(self, user_info):
        """显示主系统窗口（需要实现）"""
        # TODO: 实现跳转到主界面的逻辑

        show_QMessageBox("登录成功", f"欢迎，{user_info['username']}！", QMessageBox.Information)
        from main import MainWindow as MainSystemWindow
        self.main_window = MainSystemWindow(user_info=user_info)
        self.main_window.show()
        self.close()


if __name__ == "__main__":
    # 在创建QApplication之前设置Qt.AA_ShareOpenGLContexts属性
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
    QApplication.setAttribute(Qt.AA_ShareOpenGLContexts)
    
    app = QApplication(sys.argv)
    window = LoginMainWindow()
    window.show()
    sys.exit(app.exec_())