import os
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from message_styles import MESSAGE_BOX_STYLE

def show_QMessageBox(title, text, icon=QMessageBox.Information, buttons=None ):
    msg = QMessageBox()
    msg.setIcon(icon)
    msg.setWindowTitle(title)
    msg.setText(text)
    if buttons:
        msg.setStandardButtons(buttons)
    else:
        msg.setStandardButtons(QMessageBox.Ok)
    # QMessageBox.information(self, "登录成功", f"欢迎，{user_info['username']}！")
    msg.setStyleSheet(MESSAGE_BOX_STYLE)
    return msg.exec_()        