"""
消息框样式定义
专门用于QMessageBox的样式配置
"""

# 消息框基础样式
MESSAGE_BOX_STYLE = """
QMessageBox {
    background-color: #E6F7E6;
    color: #2D5F2D;
    font-family: "Microsoft YaHei", Arial, sans-serif;
}
QMessageBox QLabel {
    color: #2D5F2D;
    font-size: 14px;
    background-color: transparent;
}
QMessageBox QPushButton {
    background-color: #2E8B57;
    color: #ffffff;
    border: 1px solid #256D46;
    border-radius: 4px;
    padding: 8px 16px;
    font-size: 12px;
    min-width: 80px;
}
QMessageBox QPushButton:hover {
    background-color: #256D46;
    border-color: #1F5C3B;
}
QMessageBox QPushButton:pressed {
    background-color: #1F5C3B;
}
QMessageBox QPushButton:focus {
    outline: none;
}
QDialogButtonBox {
    alignment: center;
    qproperty-centerButtons: true;
}
"""