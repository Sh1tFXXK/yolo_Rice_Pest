# -*- coding: utf-8 -*-
import sys
import os
# 在导入QtWebEngineWidgets之前设置Qt.AA_ShareOpenGLContexts属性
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
from PIL import Image, ImageDraw, ImageFont
QApplication.setAttribute(Qt.AA_ShareOpenGLContexts)

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QStringListModel
from config_text import MAIN_TEXTS
from ultralytics import YOLO
from collections import Counter
from ModifyUserDialog import ModifyUserDialog
from QMessageBox_helper import show_QMessageBox
import cv2
import numpy as np
import time
import json
import torch

# ====================== 新增1：害虫介绍及防治方法知识库 ======================
PEST_KNOWLEDGE = {
    "褐飞虱": """【害虫介绍】
褐飞虱是水稻头号毁灭性害虫，成虫、若虫群集稻株基部刺吸汁液，造成稻株倒伏、枯死（俗称“冒穿”），还会传播水稻条纹叶枯病、黑条矮缩病等病毒病。
【防治方法】
1. 农业防治：选用抗虫品种；合理密植，避免偏施氮肥；适时烤田，降低田间湿度。
2. 生物防治：保护青蛙、蜘蛛、稻虱缨小蜂等天敌；放养鸭群捕食害虫。
3. 化学防治：若虫盛发期用吡虫啉、噻虫嗪、吡蚜酮等药剂，对准稻株基部喷雾。""",
    "​绿叶蝉​": """【害虫介绍】
绿叶蝉（小绿叶蝉）是水稻常见刺吸式害虫，成虫、若虫群集于水稻叶片背面刺吸汁液，导致叶片失绿、卷曲、变黄，严重时整株枯死，还会传播水稻病毒病，影响光合作用和产量。
【防治方法】
1. 农业防治：清除田间杂草，减少虫源；合理施肥，避免植株徒长；浅水勤灌，降低田间湿度。
2. 生物防治：保护瓢虫、草蛉、蜘蛛、寄生蜂等天敌，利用生物链控制虫口数量。
3. 化学防治：若虫高峰期选用啶虫脒、吡蚜酮、烯啶虫胺等药剂，兑水喷雾，注意轮换用药避免抗药性。""",
    "​稻纵卷叶螟​": """【害虫介绍】
稻纵卷叶螟幼虫吐丝卷叶，啃食叶肉形成白色条斑，严重时全田叶片发白，破坏光合作用，导致水稻减产10%-30%。
【防治方法】
1. 农业防治：合理施肥，避免植株嫩绿徒长；及时清除田间残株，减少越冬虫源。
2. 生物防治：释放赤眼蜂寄生虫卵；喷施苏云金杆菌（Bt）制剂杀灭幼虫。
3. 化学防治：幼虫孵化盛期用氯虫苯甲酰胺、甲维盐等药剂喷雾，重点喷洒叶片正面和背面。""",
    "​稻蝽​": """【害虫介绍】
稻蝽以成虫、若虫刺吸水稻茎秆、叶鞘和稻穗汁液，造成谷粒不实、空瘪，严重时稻株枯黄，影响结实率和千粒重。
【防治方法】
1. 农业防治：清除田间及周边杂草，破坏越冬场所；合理密植，改善田间通风透光条件。
2. 生物防治：保护蜘蛛、青蛙、寄生蜂等天敌。
3. 化学防治：若虫盛发期用噻虫嗪、氯虫苯甲酰胺等药剂喷雾，重点喷洒稻穗和茎秆部位。""",
    "稻螟虫": """【害虫介绍】
稻螟虫（二化螟、三化螟）幼虫蛀食水稻茎秆，造成枯心苗、白穗，严重时导致水稻大面积减产，是水稻钻蛀性核心害虫。
【防治方法】
1. 农业防治：深耕翻土，消灭越冬虫源；合理安排播期，避开螟虫发生高峰期；及时清除枯心苗。
2. 生物防治：释放赤眼蜂；施用杀螟杆菌等生物农药。
3. 化学防治：卵孵化盛期至幼虫钻蛀前，用氯虫苯甲酰胺、三唑磷等药剂喷雾。""",
    "稻秆潜蝇​": """【害虫介绍】
稻秆潜蝇幼虫潜入水稻茎秆内取食，造成心叶扭曲、枯白，分蘖减少，植株矮化，严重影响水稻生长发育。
【防治方法】
1. 农业防治：清除田间杂草，减少虫源；合理施肥，增强植株抗虫性；适时晒田，降低田间湿度。
2. 生物防治：保护寄生蜂等天敌。
3. 化学防治：幼虫孵化期用阿维菌素、吡虫啉等药剂喷雾，重点喷洒心叶部位。"""
}
# ============================================================================

# 全局类别名称映射
CLASS_NAME_MAP = {'brown-planthopper' : "褐飞虱", 'green-leafhopper' : "​绿叶蝉​", 'leaf-folder' : "​稻纵卷叶螟", 'rice-bug' : "​稻蝽", 'stem-borer' : "稻螟虫", 'whorl-maggot' : "稻秆潜蝇​"}

# 导入自动生成的界面类
from main_layout import Ui_MainWindow

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, user_info=None):
        super().__init__()
        self.user_info = user_info
        # 初始化UI界面
        self.setupUi(self)
        self.setWindowTitle(MAIN_TEXTS["title"])

        self.image_label.setVisible(False)
        self.image_bg.setUrl(QUrl.fromLocalFile("/main_img.html"))

        # ====================== 新增2：检测控制核心属性 ======================
        self.current_media_type = None  # 媒体类型：None/image/video/folder/camera
        self.current_media_data = None  # 存储加载的媒体数据
        self.is_detecting = False       # 检测状态标记，防止重复点击
        # ======================================================================

        self.init_ui()
        self.all_detections = []
        self.filtered_detections = []
        self.collapsed = False

    def update_pixmap(self, image, pixmap,size=None):
        """
        更新并显示pixmap。
        
        如果pixmap为空，则清空标签。否则，将其按比例缩放以适应标签大小，并设置。
        """
        if pixmap.isNull():
            image.clear()
            return
        if size is not None:
            scaled_pixmap = pixmap.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        else:
            print(image.size())
            scaled_pixmap = pixmap.scaled(image.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        image.setPixmap(scaled_pixmap)

    def init_ui(self):
        """
        初始化主内容区域的用户界面。
        
        按顺序创建并添加参数区、状态栏、图像区和结果表格。
        """
        # 按钮列表（图片路径+文字）
        self.button_defs = [
            ("个人中心", "icon/user.png"),
            ("图片选择", "icon/image.png"),
            ("视频选择", "icon/video.png"),
            ("文件夹选择", "icon/folder.png"),
            ("摄像头", "icon/camera.png"),
            ("模型选择", "icon/robot.png")
        ]

        self.init_Left_widget()
        self.init_table_widget()
        # 加载模型
        self.model = YOLO('runs/train_yolo12n/weights/best.pt')

        self.chinese_names = {k: CLASS_NAME_MAP.get(v, v) for k, v in self.model.names.items()}

        logo_path = 'icon/logo.png'
        logo_pixmap = QPixmap(logo_path)
        size = QSize(100,100)
        self.update_pixmap(self.logoLabel,logo_pixmap,size)  # 使用set_pixmap方法设置图片

        # 参数调节区
        self.create_param_widget()

        # 过滤器
        self.create_filter_widget()

        self.init_listview()

        # 根据是否有用户信息，决定是否显示个人中心
        if not self.user_info:
            self.btnUser2.setVisible(False)
            self.btnUser.setVisible(False)
        else:
            self.btnUser2.setVisible(True)
            self.btnUser.setVisible(True)

        self.hidden_user_info()
        self.init_user()

        self.set_shadow(self.statsGroup)
        self.set_shadow(self.filterGroup)
        self.set_shadow(self.classGroup)
        self.set_shadow(self.confGroup)
        self.set_shadow(self.posGroup)
        self.set_shadow(self.avatarWidget)

        # ====================== 新增3：连接表格行选择信号 ======================
        self.table.clicked.connect(self.on_detection_selected)

    # ====================== 新增4：开始检测核心槽函数 ======================
    def start_detection(self):
        """点击开始检测按钮执行的核心检测逻辑"""
        # 边界判断：未加载媒体
        if self.current_media_type is None or self.current_media_data is None:
            show_QMessageBox("提示", "请先选择图片/视频/文件夹/摄像头，再点击开始检测", QMessageBox.Warning)
            return
        
        # 边界判断：正在检测中，防止重复点击
        if self.is_detecting:
            # 如果是视频/摄像头，点击则停止检测
            if self.current_media_type in ["video", "camera"]:
                self.close_cap()
                self.btnStartDetect.setText("开始检测")
                self.is_detecting = False
                self.btnStartDetect.setEnabled(True)
            return

        # 标记检测状态
        self.is_detecting = True
        self.btnStartDetect.setEnabled(False)
        self.btnStartDetect.setText("检测中...")
        self.set_default_filter()
        QApplication.processEvents()

        try:
            # ========== 1. 图片检测 ==========
            if self.current_media_type == "image":
                self.original_cv_img = self.current_media_data
                results, detection_time = self.perform_detection(self.original_cv_img)
                self.process_detection_results(detection_time, results, self.current_image_path)

            # ========== 2. 文件夹批量检测 ==========
            elif self.current_media_type == "folder":
                image_files = self.current_media_data
                for image_path in image_files:
                    if not self.is_detecting:  # 支持中途停止
                        break
                    try:
                        self.original_cv_img = self.cv_imread(image_path)
                        self.current_image_path = image_path
                        if self.original_cv_img is None:
                            print(f"无法加载图片: {image_path}")
                            continue
                        
                        results, detection_time = self.perform_detection(self.original_cv_img)
                        self.process_detection_results(detection_time, results, image_path)
                        QApplication.processEvents()
                        time.sleep(1)
                    except Exception as e:
                        print(f"处理图片时出错: {image_path}, 错误: {str(e)}")
                        continue

            # ========== 3. 视频检测 ==========
            elif self.current_media_type == "video":
                self.cap = self.current_media_data
                if not self.cap.isOpened():
                    show_QMessageBox("错误", "视频文件加载失败，请重新选择", QMessageBox.Critical)
                    self.reset_detect_state()
                    return
                
                self.timer = QTimer(self)
                self.timer.timeout.connect(self.update_frame)
                fps = self.cap.get(cv2.CAP_PROP_FPS)
                interval = int(1000 / fps)
                self.timer.start(interval)
                self.btnStartDetect.setText("停止检测")
                self.btnStartDetect.setEnabled(True)
                return

            # ========== 4. 摄像头实时检测 ==========
            elif self.current_media_type == "camera":
                self.cap = self.current_media_data
                if not self.cap.isOpened():
                    show_QMessageBox("错误", "摄像头打开失败，请检查设备", QMessageBox.Critical)
                    self.reset_detect_state()
                    return
                
                self.timer = QTimer(self)
                self.timer.timeout.connect(self.update_frame)
                self.timer.start(30)
                self.btnStartDetect.setText("停止检测")
                self.btnStartDetect.setEnabled(True)
                return

        except Exception as e:
            print(f"检测过程出错: {str(e)}")
            show_QMessageBox("检测失败", f"检测过程中出现错误: {str(e)}", QMessageBox.Critical)
        
        # 检测完成，重置状态
        self.reset_detect_state()

    # ====================== 新增5：重置检测状态辅助函数 ======================
    def reset_detect_state(self):
        """重置检测按钮和状态"""
        self.is_detecting = False
        self.btnStartDetect.setText("开始检测")
        self.btnStartDetect.setEnabled(True)

    # ====================== 新增6：清空检测结果辅助函数 ======================
    def clear_detect_result(self):
        """清空界面上的检测结果，加载新媒体时调用"""
        self.all_detections = []
        self.filtered_detections = []
        self.table.setRowCount(0)
        self.time_label.setText("检测耗时: 0.00s")
        self.count_label.setText("检测目标: 0个")
        self.listmodel.setStringList([])
        self.category_value.setText("未选择")
        self.confidence_value.setText("0.00")
        self.xmin_value.setText("0")
        self.ymin_value.setText("0")
        self.xmax_value.setText("0")
        self.ymax_value.setText("0")

    def init_user(self):
         # 根据是否有用户信息，决定是否显示个人中心
        if not self.user_info:
            self.btnUser2.setVisible(False)
            self.btnUser.setVisible(False)
        else:
            self.btnUser2.setVisible(True)
            self.btnUser.setVisible(True)
            username = self.user_info.get('username', '未知用户') if self.user_info else '未知用户'
            self.username_value.setText(username)
            avatar_path = self.user_info.get('avatar', 'avator/default.jpg') if self.user_info else 'avator/default.jpg'
            avatar_path_for_stylesheet = avatar_path.replace('\\', '/')
            avatar_size = 180  # 将头像尺寸增加到200px
            self.avatar_label.setStyleSheet(f"""
                QLabel {{
                    border: 5px solid #2980b9;
                    border-radius: {avatar_size / 2}px; /* 确保是完美的圆形 */
                    border-image: url({avatar_path_for_stylesheet}) 0 0 0 0 stretch stretch;
                }}
            """)

    def set_shadow(self, widget,r=79,g=172,b=254,a=100):
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(r, g, b, a))
        shadow.setOffset(0, 2)
        widget.setGraphicsEffect(shadow)

    def init_listview(self):
        self.listmodel = QStringListModel()
        self.catlistView.setModel(self.listmodel) 

    def init_Left_widget(self):   
        self.collapse_widget2.setVisible(False)
        self.collapse_widget.setVisible(True)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setText(MAIN_TEXTS["panel_title"])

        # ====================== 新增7：创建差异化的开始检测按钮 ======================
        # 1. 获取左侧按钮的垂直布局（兼容Qt Designer生成的UI）
        left_layout = self.collapse_widget.layout()
        if not left_layout:
            left_layout = QVBoxLayout(self.collapse_widget)
            self.collapse_widget.setLayout(left_layout)

        # 2. 添加间距（和模型选择按钮空开距离）
        left_layout.addSpacing(30)

        # 3. 创建开始检测按钮
        self.btnStartDetect = QPushButton("开始检测")
        self.btnStartDetect.setFixedHeight(55)  # 比其他按钮更高，更醒目
        self.btnStartDetect.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))  # 加粗+更大字号
        self.btnStartDetect.setEnabled(False)  # 初始禁用，加载媒体后启用

        # 4. 差异化样式：红橙渐变，和其他按钮完全区分，带hover效果
        self.btnStartDetect.setStyleSheet("""
            QPushButton {
                border-radius: 12px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff4444, stop:1 #ff8800);
                color: white;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff2222, stop:1 #ff7700);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #dd2222, stop:1 #dd6600);
                padding-left: 3px;
                padding-top: 3px;
            }
            QPushButton:disabled {
                background: #cccccc;
                color: #888888;
            }
        """)

        # 5. 添加按钮到布局，连接点击事件
        left_layout.addWidget(self.btnStartDetect)
        self.btnStartDetect.clicked.connect(self.start_detection)
        # ======================================================================

    # ====================== 修改1：重构表格初始化方法 ======================
    def init_table_widget(self):
        # 设置表格样式
        row_count = 0
        self.table.setRowCount(row_count)
        for row in range(row_count):
            for col in range(self.table.columnCount()):
                self.table.setItem(row, col, QTableWidgetItem(""))  # 空内容
        
        # 设置表头（核心：替换“文件路径”为“害虫介绍及防治”）
        self.table.setHorizontalHeaderLabels(["序号", "害虫介绍及防治", "类别", "置信度", "坐标位置"])
        
        # 设置列宽和自适应
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 防治介绍列自适应拉伸
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        
        self.table.setColumnWidth(0, 80)    # 序号列
        self.table.setColumnWidth(2, 120)   # 类别列
        self.table.setColumnWidth(3, 100)   # 置信度列
        self.table.setColumnWidth(4, 250)   # 坐标列
        
        # 关键优化：开启自动换行+行高自适应
        self.table.setWordWrap(True)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setFixedHeight(500)

    def update_user_display(self, updated_info):
        """
        槽函数：接收用户信息更新信号并刷新UI。
        
        主要用于在修改头像后，实时更新右侧面板的头像显示。

        Args:
            updated_info (dict): 包含更新信息的字典，如 {'avatar': 'new_path'}。
        """
        if 'avatar' in updated_info and self.user_info:
            self.user_info['avatar'] = updated_info['avatar']
            self.init_user()
            print(f"[DEBUG] 主界面头像已更新为: {self.user_info['avatar']}")

    def prompt_logout(self):
        """
        弹出自定义的注销确认对话框。
        
        如果用户确认，则发射 `logout_requested` 信号。
        """
        reply = show_QMessageBox("注销确认", "您确定要注销并返回到登录界面吗？", QMessageBox.Question, QMessageBox.Yes | QMessageBox.Cancel)
        if reply == QMessageBox.Yes:
            from login import LoginMainWindow
            self.login_window = LoginMainWindow()
            self.login_window.show()
            self.close()

    def open_modify_user_dialog(self):
        """
        打开修改用户信息对话框。
        
        创建并显示 `ModifyUserDialog`，并连接其 `user_updated` 信号到本控件的刷新槽函数。
        """
        if self.user_info:
            dialog = ModifyUserDialog(self.user_info['username'], self)
            dialog.user_updated.connect(self.update_user_display)
            dialog.exec_()

    def show_user_info(self):
        """
        切换到并显示用户信息视图。
        
        此视图包含用户头像、用户名和操作按钮（如修改信息、注销）。
        """
        print(f"[DEBUG] RightPanelWidget.show_user_info() called. user_info is: {self.user_info}")
        self.rightWidget.setVisible(False)
        self.userWidget.setVisible(True)

    def hidden_user_info(self):
        self.rightWidget.setVisible(True)
        self.userWidget.setVisible(False)

    def confValueChanged(self, v):
        print(f"置信度 (Conf): {v/100:.2f}")
        self.conf_label.setText(f"置信度 (Conf): {v/100:.2f}")

    def iouValueChanged(self, v):
        print(f"交并比 (IoU): {v/100:.2f}")
        self.iou_label.setText(f"交并比 (IoU): {v/100:.2f}")

    def create_param_widget(self):
        self.conf_label.setText("置信度 (Conf): 0.5")
        self.iou_label.setText("交并比 (IoU): 0.5")

    def create_filter_widget(self):
        # 先添加“显示所有类别”选项
        self.filter_combo.addItem("显示所有类别")
        self.filter_combo.addItems(list(CLASS_NAME_MAP.values()))    
        # 连接过滤信号
        self.filter_combo.currentTextChanged.connect(self.on_filter_changed)

    def on_filter_changed(self, selected_text):
        """
        处理类别过滤下拉框的变化事件。
        """
        print("on_filter_changed:", selected_text)
        self.filter_detections_by_category(selected_text)

    def toggle_collapse(self):
        """
        切换侧边栏的展开和收缩状态。
        """
        self.collapsed = not self.collapsed
        if self.collapsed:
            self.collapse_widget2.setVisible(True)
            self.collapse_widget.setVisible(False)
        else:
            self.collapse_widget2.setVisible(False)
            self.collapse_widget.setVisible(True)

    def select_model_file(self):
        """
        响应“模型选择”按钮点击事件。
        """
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,
                                                  "选择YOLO模型权重文件",
                                                  "",
                                                  "PyTorch模型文件 (*.pt);;所有文件 (*)",
                                                  options=options)
        if fileName:
            try:
                print(f"选择的模型文件: {fileName}")
                new_model = YOLO(fileName)
                self.model = new_model
                self.chinese_names = {k: CLASS_NAME_MAP.get(v, v) for k, v in self.model.names.items()}
                show_QMessageBox("模型加载成功", f"模型已成功加载！\n\n文件路径: {fileName}")
                print(f"✅ 模型加载成功: {fileName}")
            except Exception as e:
                print(f"❌ 模型加载失败: {str(e)}")
                show_QMessageBox("模型加载失败", f"模型加载失败！\n\n错误信息: {str(e)}\n\n请检查文件是否为有效的YOLO模型文件。", QMessageBox.Critical)

    def cv_imread(self, filePath):
        """
        使用OpenCV读取可能包含中文路径的图片。
        """
        cv_img = cv2.imdecode(np.fromfile(filePath, dtype=np.uint8), -1)
        return cv_img

    def filter_detections_by_category(self, selected_category):
        """
        根据右侧面板下拉框选择的类别，过滤并重新绘制检测结果。
        """
        if not hasattr(self, 'all_detections') or not self.all_detections :
            return
        if selected_category == "显示所有类别":
            self.filtered_detections = self.all_detections
        else:
            self.filtered_detections = [det for det in self.all_detections if det['name'] == selected_category]
        self.draw_filtered_detections(self.filtered_detections)
        self.update_table_with_filtered_detections(self.filtered_detections)    

    def draw_filtered_detections(self, filtered_detections=None,selected_index=-1):
        """
        在主图像上仅绘制过滤后的检测框和标签。
        """
        if self.original_cv_img is None:
            return

        # 1. 准备字体
        font = None
        try:
            font_paths = [
                '/System/Library/Fonts/PingFang.ttc',
                '/System/Library/Fonts/STHeiti Medium.ttc',
                'C:/Windows/Fonts/msyh.ttc',
                'C:/Windows/Fonts/simhei.ttf',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'
            ]
            for path in font_paths:
                if os.path.exists(path):
                    font = ImageFont.truetype(path, 20)
                    break
            if font is None:
                font = ImageFont.load_default()
        except Exception as e:
            print(f"加载字体时出错: {str(e)}")
            font = ImageFont.load_default()

        # 2. 准备图像副本和颜色
        result_image = self.original_cv_img.copy()
        overlay = result_image.copy()
        colors = {0: (0, 255, 0), 1: (0, 0, 255), 2: (0, 255, 255), 3: (255, 0, 255), 4: (255, 165, 0), 5: (128, 0, 128), 6: (0, 128, 128), 
                  7: (255, 192, 203), 8: (64, 224, 208), 9: (240, 230, 140), 10: (255, 215, 0), 11: (138, 43, 226), 
                  12: (220, 20, 60), 13: (0, 191, 255), 14: (50, 205, 50), 15: (255, 140, 0), 16: (199, 21, 133), 
                  17: (70, 130, 180), 18: (210, 105, 30), 19: (100, 149, 237)}
        default_color = (255, 0, 0)

        # 3. 绘制过滤后的检测框
        for i,det in enumerate(filtered_detections):
            if selected_index != -1 and i != selected_index:
                continue
            box = det['box']
            x1, y1, x2, y2 = int(box['x1']), int(box['y1']), int(box['x2']), int(box['y2'])
            cls_id = det['class']
            
            color_bgr = colors.get(cls_id, default_color)
            thickness = 3

            # 绘制边界框
            cv2.rectangle(result_image, (x1, y1), (x2, y2), color_bgr, thickness)

            # 准备标签文本
            name = det['name']
            conf = det['confidence']
            label = f"{name} {conf:.2f}"
            
            # 计算文本尺寸
            text_bbox = ImageDraw.Draw(Image.new('RGB', (0,0))).textbbox((0, 0), label, font=font)
            text_width, text_height = text_bbox[2] - text_bbox[0], text_bbox[3] - text_bbox[1]

            # 在覆盖层上绘制标签背景
            cv2.rectangle(overlay, (x1, y1 - text_height - 10), (x1 + text_width + 8, y1), color_bgr, -1)

        # 4. 合并覆盖层，实现半透明效果
        result_image = cv2.addWeighted(overlay, 0.6, result_image, 0.4, 0)

        # 5. 转换为PIL，绘制文本
        pil_image = Image.fromarray(cv2.cvtColor(result_image, cv2.COLOR_BGR2RGB))
        draw = ImageDraw.Draw(pil_image)
        for i,det in enumerate(filtered_detections):
            if selected_index != -1 and i != selected_index:
                continue
            box = det['box']
            x1, y1 = int(box['x1']), int(box['y1'])
            name = det['name']
            conf = det['confidence']
            label = f"{name} {conf:.2f}"
            
            try:
                text_bbox = draw.textbbox((0, 0), label, font=font)
                text_height = text_bbox[3] - text_bbox[1]
            except AttributeError:
                _, text_height = draw.textsize(label, font=font)
                
            draw.text((x1 + 5, y1 - text_height - 7), label, font=font, fill=(255, 255, 255))
        
        # 6. 转回OpenCV格式并显示
        final_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        img_rgb = cv2.cvtColor(final_image, cv2.COLOR_BGR2RGB)
        height, width, channel = img_rgb.shape
        bytesPerLine = 3 * width
        qImg = QImage(img_rgb.data, width, height, bytesPerLine, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(qImg)
        self.update_pixmap(self.image_label, pixmap)

    # ====================== 修改2：重构表格数据填充方法 ======================
    def update_table_with_filtered_detections(self, filtered_detections):
        """
        使用过滤后的检测结果列表更新结果表格（核心：替换文件路径为害虫介绍）。
        """
        self.table.setRowCount(0)  # 清空表格
        for i, det in enumerate(filtered_detections):
            name = det['name']
            conf = det['confidence']
            box = det['box']
            x1, y1, x2, y2 = int(box['x1']), int(box['y1']), int(box['x2']), int(box['y2'])

            self.table.insertRow(i)
            # 1. 序号 - 居中
            seq_item = QTableWidgetItem(str(i + 1))
            seq_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 0, seq_item)
            
            # 2. 核心：害虫介绍及防治（替换原文件路径）
            pest_info = PEST_KNOWLEDGE.get(name, "⚠️ 暂无该害虫的介绍与防治方法")
            info_item = QTableWidgetItem(pest_info)
            info_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)  # 文本左上对齐，方便阅读
            info_item.setBackground(QColor("#f5f5f5"))  # 浅灰色背景，提升可读性
            self.table.setItem(i, 1, info_item)
            
            # 3. 类别 - 居中
            class_item = QTableWidgetItem(name)
            class_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 2, class_item)
            
            # 4. 置信度 - 居中
            conf_item = QTableWidgetItem(f"{conf:.2f}")
            conf_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 3, conf_item)
            
            # 5. 坐标 - 居中
            coord_item = QTableWidgetItem(f"({x1},{y1},{x2},{y2})")
            coord_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 4, coord_item)
        
        # 关键：自动调整行高以适配长文本
        self.table.resizeRowsToContents()

    def on_detection_selected(self):
        """
        处理结果表格中的行选择事件。
        """
        selected_row = self.table.currentRow()
        if selected_row < 0 or selected_row >= len(self.filtered_detections):
            return

        # 高亮显示选中的检测框
        self.draw_filtered_detections(filtered_detections=self.filtered_detections,selected_index=selected_row)

        # 更新右侧面板信息
        det = self.filtered_detections[selected_row]
        name = det['name']
        conf = det['confidence']
        box = det['box']
        x1, y1, x2, y2 = int(box['x1']), int(box['y1']), int(box['x2']), int(box['y2'])
        
        self.category_value.setText(name)
        self.confidence_value.setText(f"{conf:.2f}")
        self.xmin_value.setText(str(x1))
        self.ymin_value.setText(str(y1))
        self.xmax_value.setText(str(x2))
        self.ymax_value.setText(str(y2))

    def perform_detection(self, image):
        """
        执行目标检测的核心逻辑。
        """
        start_time = time.time()
        conf_threshold = self.conf_slider.value() / 100.0
        iou_threshold = self.iou_slider.value() / 100.0
        results = self.model.predict(source=image, conf=conf_threshold, iou=iou_threshold)
        detection_time = time.time() - start_time
        return results, detection_time

    def process_detection_results(self, detection_time, results, image_path):
        """
        处理检测结果的通用函数。
        """
        self.image_bg.setVisible(False)
        self.image_label.setVisible(True)
        try:
            # 1. 解析并保存检测结果
            detections_json = json.loads(results[0].to_json())
            self.all_detections = []
            for det in detections_json:
                english_name = det['name']
                det['name'] = CLASS_NAME_MAP.get(english_name, english_name)  # 翻译为中文
                self.all_detections.append(det)
            
            selected_category = self.filter_combo.currentText()
            if selected_category == "显示所有类别":
                self.filtered_detections = self.all_detections
            else:
                self.filtered_detections = [det for det in self.all_detections if det['name'] == selected_category]

            # 2. 更新UI
            self.time_label.setText(f"检测耗时: {detection_time:.2f}s")
            self.count_label.setText(f"检测目标: {len(self.filtered_detections)}个")

            name_counts = Counter(item['name'] for item in self.filtered_detections)
            result = [f"{name}: {count}" for name, count in name_counts.items()]
            self.listmodel.setStringList(result)
            self.catlistView.setModel(self.listmodel) 

            # 3. 填充表格
            self.update_table_with_filtered_detections(self.filtered_detections)

            # 4. 绘制所有检测框
            self.draw_filtered_detections(filtered_detections=self.filtered_detections)

            # 5. 更新右侧面板
            self.category_value.setText("未选择")
            self.confidence_value.setText("0.00")
            self.xmin_value.setText("0")
            self.ymin_value.setText("0")
            self.xmax_value.setText("0")
            self.ymax_value.setText("0")
                
            return True
        except Exception as e:
            print(f"处理检测结果时出错: {str(e)}")
            return False
    
    def set_default_filter(self):
        self.filter_combo.blockSignals(True)
        self.filter_combo.setCurrentIndex(0)  
        self.filter_combo.blockSignals(False)  

    # ====================== 修改3：图片选择-仅加载不检测 ======================
    def open_image_file(self):
        """响应“图片选择”按钮点击事件：仅加载图片，不执行检测"""
        self.hidden_user_info()
        self.close_cap()
        self.reset_detect_state()
        
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,
                                                  "选择一张图片",
                                                  "",
                                                  "图片文件 (*.png *.jpg *.jpeg *.bmp);;所有文件 (*)",
                                                  options=options)
        if fileName:
            try:
                print(f"选择的图片文件: {fileName}")
                cv_img = self.cv_imread(fileName)
                self.current_image_path = fileName
                if cv_img is None:
                    show_QMessageBox("加载失败", "无法加载所选的图片文件。", QMessageBox.Warning)
                    return

                # 仅加载图片到界面，不执行检测
                self.original_cv_img = cv_img
                self.image_bg.setVisible(False)
                self.image_label.setVisible(True)
                # 显示原始图片
                img_rgb = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB)
                height, width, channel = img_rgb.shape
                bytesPerLine = 3 * width
                qImg = QImage(img_rgb.data, width, height, bytesPerLine, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(qImg)
                self.update_pixmap(self.image_label, pixmap)

                # 保存媒体数据，启用检测按钮
                self.current_media_type = "image"
                self.current_media_data = cv_img
                self.btnStartDetect.setEnabled(True)

                # 清空上一次的检测结果
                self.clear_detect_result()

            except Exception as e:
                print(f"图片加载过程中出现错误: {str(e)}")
                show_QMessageBox("加载失败", f"图片加载过程中出现错误: {str(e)}", QMessageBox.Warning)

    # ====================== 修改4：文件夹选择-仅加载不检测 ======================
    def open_folder(self):
        """响应“文件夹选择”按钮点击事件：仅加载图片列表，不执行检测"""
        self.hidden_user_info()
        self.close_cap()
        self.reset_detect_state()

        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "选择包含图片的文件夹", "", options=options)
        
        if folder_path:
            image_extensions = ['.jpg', '.jpeg', '.png', '.bmp']
            image_files = []
            
            for file in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file)
                if os.path.isfile(file_path) and any(file.lower().endswith(ext) for ext in image_extensions):
                    image_files.append(file_path)
            
            if not image_files:
                show_QMessageBox("没有图片", "所选文件夹中没有找到支持的图片文件。", QMessageBox.Warning)
                return
            
            # 保存媒体数据，启用检测按钮
            self.current_media_type = "folder"
            self.current_media_data = sorted(image_files)
            self.btnStartDetect.setEnabled(True)

            # 清空上一次的检测结果，更新界面提示
            self.clear_detect_result()
            self.image_bg.setVisible(True)
            self.image_label.setVisible(False)
            show_QMessageBox("加载成功", f"已加载文件夹，共找到 {len(image_files)} 张图片\n点击「开始检测」执行批量检测", QMessageBox.Information)

    def close_cap(self):
        if hasattr(self, 'timer') and self.timer.isActive():
                print("[DEBUG] 停止正在播放的内容...")
                self.timer.stop()
                if hasattr(self, 'cap') and self.cap.isOpened():
                    self.cap.release()
                    print("[DEBUG] 资源已释放")
        self.reset_detect_state()

    # ====================== 修改5：视频选择-仅加载不检测 ======================
    def open_video_file(self):
        """响应“视频选择”按钮点击事件：仅加载视频文件，不启动检测"""
        self.hidden_user_info()
        self.close_cap()
        self.reset_detect_state()

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,
                                                  "选择一个视频文件",
                                                  "",
                                                  "视频文件 (*.mp4 *.avi *.mov *.mkv);;所有文件 (*)",
                                                  options=options)
        if fileName:
            cap = cv2.VideoCapture(fileName)
            if not cap.isOpened():
                show_QMessageBox("加载失败", "无法加载所选的视频文件。", QMessageBox.Warning)
                return
            
            # 仅加载视频，不启动定时器
            self.current_media_type = "video"
            self.current_media_data = cap
            self.btnStartDetect.setEnabled(True)

            # 清空上一次的检测结果，显示视频第一帧
            self.clear_detect_result()
            ret, first_frame = cap.read()
            if ret:
                self.original_cv_img = first_frame
                self.image_bg.setVisible(False)
                self.image_label.setVisible(True)
                img_rgb = cv2.cvtColor(first_frame, cv2.COLOR_BGR2RGB)
                height, width, channel = img_rgb.shape
                bytesPerLine = 3 * width
                qImg = QImage(img_rgb.data, width, height, bytesPerLine, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(qImg)
                self.update_pixmap(self.image_label, pixmap)
            
            show_QMessageBox("加载成功", "视频文件已加载\n点击「开始检测」播放并执行检测", QMessageBox.Information)

    # ====================== 修改6：摄像头选择-仅打开不检测 ======================
    def open_camera(self):
        """响应“摄像头”按钮点击事件：仅打开摄像头，不启动实时检测"""
        print("[DEBUG] 开始打开摄像头...")
        self.hidden_user_info()
        self.close_cap()
        self.reset_detect_state()

        try:
            print("[DEBUG] 系统信息:")
            print(f"[DEBUG] 操作系统: {sys.platform}")
            print(f"[DEBUG] Python版本: {sys.version}")
            print(f"[DEBUG] OpenCV版本: {cv2.__version__}")
            
            print("[DEBUG] 尝试打开默认摄像头(索引0)...")
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                print("[ERROR] 无法打开默认摄像头")
                print("[DEBUG] 尝试打开备用摄像头(索引1)...")
                cap = cv2.VideoCapture(1)
                if not cap.isOpened():
                    print("[ERROR] 备用摄像头也无法打开")
                    show_QMessageBox("摄像头错误", "无法打开摄像头，请检查设备连接。详细信息请查看控制台日志。", QMessageBox.Warning)
                    return
            
            width = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            height = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            fps = cap.get(cv2.CAP_PROP_FPS)
            print(f"[DEBUG] 摄像头成功打开! 分辨率: {width}x{height}, FPS: {fps}")
            
        except Exception as e:
            print(f"[ERROR] 打开摄像头时发生异常: {str(e)}")
            show_QMessageBox("摄像头错误", f"打开摄像头时发生异常: {str(e)}", QMessageBox.Warning)
            return

        # 仅打开摄像头，不启动定时器
        self.current_media_type = "camera"
        self.current_media_data = cap
        self.btnStartDetect.setEnabled(True)

        # 清空上一次的检测结果，显示摄像头第一帧
        self.clear_detect_result()
        ret, first_frame = cap.read()
        if ret:
            self.original_cv_img = first_frame
            self.image_bg.setVisible(False)
            self.image_label.setVisible(True)
            img_rgb = cv2.cvtColor(first_frame, cv2.COLOR_BGR2RGB)
            height, width, channel = img_rgb.shape
            bytesPerLine = 3 * width
            qImg = QImage(img_rgb.data, width, height, bytesPerLine, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qImg)
            self.update_pixmap(self.image_label, pixmap)
        
        show_QMessageBox("加载成功", "摄像头已打开\n点击「开始检测」启动实时检测", QMessageBox.Information)

    def update_frame(self):
        """
        由定时器调用的槽函数，用于更新一帧画面并进行检测。
        """
        try:
            ret, frame = self.cap.read()
            if not ret:
                print("[ERROR] 无法读取画面或者视频读取结束")
                self.close_cap()
                show_QMessageBox("提示", "视频播放完毕/摄像头读取中断", QMessageBox.Information)
                return

            self.original_cv_img = frame
            self.current_image_path = "实时画面"
            results, detection_time = self.perform_detection(self.original_cv_img)
            self.process_detection_results(detection_time, results, "实时画面")
            
        except Exception as e:
            print(f"[ERROR] 更新画面时发生异常: {str(e)}")
            self.close_cap()
            show_QMessageBox("错误", f"处理画面时发生异常: {str(e)}", QMessageBox.Warning)
            return

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    
    # 全局绿色水稻主题样式，强制覆盖所有界面，100%生效
    app.setStyleSheet("""
/* 全局水稻背景 + 绿色主题 */
QMainWindow {
    background-image: url(images/rice_bg.png);
    background-repeat: no-repeat;
    background-position: center;
    background-color: #E6F7E6;
}

/* 按钮 */
QPushButton {
    background-color: #2E8B57;
    color: white;
    border-radius: 10px;
    padding: 12px 20px;
    font-size: 15px;
    font-weight: 600;
    border: none;
}
QPushButton:hover {
    background-color: #256D46;
}
QPushButton:pressed {
    background-color: #1F5C3B;
}

/* 面板半透明，透出背景图 */
QFrame, QWidget, QGroupBox {
    background-color: rgba(230, 247, 230, 180);
    border-radius: 10px;
}

/* 文字 */
QLabel {
    color: #2D5F2D;
    background: transparent;
}

/* 滑块 */
QSlider::groove:horizontal {
    height: 8px;
    background: #2E8B57;
    border-radius: 4px;
}
QSlider::handle:horizontal {
    background: white;
    border: 2px solid #2E8B57;
    width: 20px;
    height: 20px;
    margin: -6px 0;
    border-radius: 10px;
}
""")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())