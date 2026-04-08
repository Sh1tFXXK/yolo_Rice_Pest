from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtCore import pyqtSignal
from PyQt5.uic import loadUiType
from QMessageBox_helper import show_QMessageBox

# 加载UI文件
UI_FILE = "modify_user_dialog.ui"
FormClass, BaseClass = loadUiType(UI_FILE)

class ModifyUserDialog(BaseClass, FormClass):
    """
    一个用于修改用户信息的对话框。
    
    允许用户修改密码和头像。当信息更新成功时，会发射 `user_updated` 信号。
    """
    # 1. 定义一个信号，当用户信息更新成功时发射
    user_updated = pyqtSignal(dict)

    def __init__(self, username, parent=None):
        """
        构造函数，初始化对话框。

        Args:
            username (str): 当前要修改信息的用户名。
            parent (QWidget, optional): 父控件。 Defaults to None.
        """
        super().__init__(parent)
        self.username = username
        self.new_avatar_path = None

        # 加载UI文件
        self.setupUi(self)
        
        # 设置窗口属性
        self.setModal(True)
        
        # 初始化UI内容
        self.init_ui()
        
        # 连接信号槽
        self.connect_signals()

    def init_ui(self):
        """初始化UI内容"""
        # 更新用户名标签
        self.usernameLabel.setText(f"👤用户名: {self.username}")
        
        # 设置头像标签的鼠标事件和光标
        self.avatarLabel.mousePressEvent = self.select_avatar
        self.avatarLabel.setCursor(Qt.PointingHandCursor)

    def connect_signals(self):
        """连接信号槽"""
        self.closeButton.clicked.connect(self.reject)
        self.cancelButton.clicked.connect(self.reject)
        self.saveButton.clicked.connect(self.save_changes)

    def select_avatar(self, event):
        """
        处理头像标签的点击事件，打开文件对话框以选择新头像。

        Args:
            event (QMouseEvent): 鼠标点击事件。
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择新头像", 
            "", 
            "图片文件 (*.png *.jpg *.jpeg)"
        )
        if file_path:
            self.new_avatar_path = file_path
            pixmap = QPixmap(file_path)
            self.avatarLabel.setPixmap(
                pixmap.scaled(
                    self.avatarLabel.size(), 
                    Qt.KeepAspectRatio, 
                    Qt.SmoothTransformation
                )
            )
            # 头像选择后，更新样式
            self.avatarLabel.setStyleSheet("""
                QLabel#avatarLabel {
                    background-color: rgba(255, 255, 255, 0.05);
                    border: 2px solid #667eea;
                    border-radius: 15px;
                }
            """)

    def save_changes(self):
        """
        保存用户所做的修改。
        
        - 校验输入的密码。
        - 调用数据库帮助函数 `update_user` 来更新信息。
        - 如果更新成功，则发射信号并关闭对话框。
        """
        password = self.passwordInput.text()
        confirm_password = self.confirmPasswordInput.text()

        if password != confirm_password:
            show_QMessageBox("错误", "两次输入的密码不一致！" ,QMessageBox.Warning)
            return

        # 如果密码字段非空但长度不足，则提示
        if password and len(password) < 6:
            show_QMessageBox("错误", "新密码长度至少需要6位！", QMessageBox.Warning)
            return

        # 只有当用户输入了新密码时，才传递密码参数
        password_to_update = password if password else None

        # 调用db_helper中的函数更新数据库
        from db_helper import update_user
        if update_user(self.username, new_password=password_to_update, new_avatar=self.new_avatar_path):
            show_QMessageBox("成功", "用户信息更新成功！", QMessageBox.Information)
            # 2. 如果头像被修改，则发射信号，携带新头像路径
            if self.new_avatar_path:
                self.user_updated.emit({'avatar': self.new_avatar_path})
            self.accept()  # 关闭对话框并返回QDialog.Accepted
        else:
            show_QMessageBox("失败", "用户信息更新失败，请稍后重试。", QMessageBox.Warning)