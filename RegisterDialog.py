from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUi
from db_helper import user_exists, add_user
from QMessageBox_helper import show_QMessageBox

class RegisterDialog(QDialog):
    """一个用于用户注册的对话框。"""
    def __init__(self, parent=None):
        """
        构造函数，初始化注册对话框的UI。

        Args:
            parent (QWidget, optional): 父控件。 Defaults to None.
        """
        super().__init__(parent)
        
        # 加载UI文件（包含所有样式）
        loadUi('register_dialog.ui', self)
        
        # 初始化变量
        self.avatar_path = "avator/default.jpg"
        
        # 设置头像预览
        self.set_avatar_preview(self.avatar_path)
        
        # 连接信号和槽
        self.setup_connections()

    def setup_connections(self):
        """连接信号和槽"""
        self.select_avatar_btn.clicked.connect(self.select_avatar)
        self.cancel_btn.clicked.connect(self.reject)
        self.register_btn.clicked.connect(self.register)

    def set_avatar_preview(self, path):
        """
        设置并显示用户选择的头像预览。

        Args:
            path (str): 头像图片的路径。
        """
        # 加载原始图片
        pixmap = QPixmap(path)
        
        # 将图片缩放到120x120像素，保持宽高比，平滑缩放
        scaled_pixmap = pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        
        # 如果图片不是正方形，创建正方形画布并居中显示
        if scaled_pixmap.width() != scaled_pixmap.height():
            square_pixmap = QPixmap(120, 120)
            square_pixmap.fill(Qt.transparent)  # 透明背景
            
            painter = QPainter(square_pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # 计算居中位置
            x = (120 - scaled_pixmap.width()) // 2
            y = (120 - scaled_pixmap.height()) // 2
            
            # 在正方形画布上绘制图片
            painter.drawPixmap(x, y, scaled_pixmap)
            painter.end()
            
            self.avatar_preview.setPixmap(square_pixmap)
        else:
            # 如果已经是正方形，直接使用
            self.avatar_preview.setPixmap(scaled_pixmap)

    def select_avatar(self):
        """打开文件对话框让用户选择头像图片。"""
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "选择头像图片", "", "Images (*.png *.xpm *.jpg *.jpeg)", options=options)
        if fileName:
            self.avatar_path = fileName
            self.set_avatar_preview(self.avatar_path)

    def register(self):
        """
        处理注册按钮的点击事件。
        
        - 校验所有输入字段。
        - 检查用户名是否已存在。
        - 调用 `db_helper.add_user` 将新用户信息存入数据库。
        - 注册成功则关闭对话框，失败则显示错误提示。
        """
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        confirm_password = self.confirm_password_input.text().strip()

        if not username or not password or not confirm_password:
            show_QMessageBox("注册失败", "请填写所有字段！", QMessageBox.Warning)
            return
        if len(username) < 3:
            show_QMessageBox("注册失败", "用户名至少需要3个字符！", QMessageBox.Warning)
            return
        if len(password) < 6:
            show_QMessageBox("注册失败", "密码至少需要6个字符！", QMessageBox.Warning)
            return
        if password != confirm_password:
            show_QMessageBox("注册失败", "两次输入的密码不一致！", QMessageBox.Warning)
            return
        if user_exists(username):
            show_QMessageBox("注册失败", "用户名已存在，请更换用户名！", QMessageBox.Warning)
            return
        if add_user(username, password, self.avatar_path):
            show_QMessageBox("注册成功", f"用户 {username} 注册成功！", QMessageBox.Information)
            self.accept()
        else:
            show_QMessageBox("注册失败", "注册失败，请重试！", QMessageBox.Warning)


# 注意：以下函数需要根据您的实际数据库实现进行替换