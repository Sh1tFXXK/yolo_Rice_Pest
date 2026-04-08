from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_working.docx")
OUT_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

W = "{%s}" % NS["w"]


def make_run(text: str):
    r = ET.Element(W + "r")
    t = ET.SubElement(r, W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def set_paragraph_text(p, text: str):
    ppr = p.find(W + "pPr")
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    p.append(make_run(text))


def set_cell_text(tbl, row_idx: int, col_idx: int, text: str):
    tr = tbl.findall(W + "tr")[row_idx - 1]
    tc = tr.findall(W + "tc")[col_idx - 1]
    paragraphs = tc.findall(W + "p")
    if not paragraphs:
        paragraphs = [ET.SubElement(tc, W + "p")]
    set_paragraph_text(paragraphs[0], text)
    for extra in paragraphs[1:]:
        tc.remove(extra)


def remove_table_row(tbl, row_idx: int):
    rows = tbl.findall(W + "tr")
    tbl.remove(rows[row_idx - 1])


def f1(p, r):
    return 2 * p * r / (p + r)


def generate_flowchart_png():
    with ZipFile(DOCX_PATH) as zf:
        original = zf.read("word/media/image3.png")
    size = Image.open(BytesIO(original)).size

    img = Image.new("RGB", size, "#f3f8ef")
    draw = ImageDraw.Draw(img)

    try:
        font_title = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 38)
        font_box = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 28)
        font_small = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 22)
    except Exception:
        font_title = ImageFont.load_default()
        font_box = ImageFont.load_default()
        font_small = ImageFont.load_default()

    w, h = size
    draw.rounded_rectangle((20, 20, w - 20, h - 20), radius=24, outline="#6f9c55", width=4, fill="#eef6e8")
    draw.text((w // 2, 60), "水稻虫害检测系统业务流程", anchor="mm", fill="#2f5f2f", font=font_title)

    boxes = [
        ("用户登录/注册", "#4c8f49"),
        ("选择检测源\n图片 / 视频 / 文件夹 / 摄像头", "#5f9f52"),
        ("加载 YOLO 模型并设置参数\n置信度 / IoU / 模型文件", "#75ad5d"),
        ("执行推理检测\n边界框、类别、置信度输出", "#8ab766"),
        ("结果展示与交互\n表格、防治说明、类别筛选", "#9fc172"),
    ]

    box_w = 620
    box_h = 120
    x1 = (w - box_w) // 2
    y = 140
    for i, (label, color) in enumerate(boxes):
        y1 = y + i * 145
        y2 = y1 + box_h
        draw.rounded_rectangle((x1, y1, x1 + box_w, y2), radius=22, fill=color, outline="#2e6031", width=3)
        draw.multiline_text((w // 2, (y1 + y2) // 2), label, anchor="mm", align="center", fill="white", font=font_box, spacing=8)
        if i < len(boxes) - 1:
            mid_x = w // 2
            start_y = y2
            end_y = y2 + 25
            draw.line((mid_x, start_y, mid_x, end_y), fill="#3e7a3e", width=6)
            draw.polygon([(mid_x - 12, end_y), (mid_x + 12, end_y), (mid_x, end_y + 18)], fill="#3e7a3e")

    note = "默认部署模型为 YOLOv12n，检测结果联动更新右侧类别、置信度与坐标信息。"
    draw.rounded_rectangle((120, h - 120, w - 120, h - 55), radius=18, fill="#dbeccf", outline="#7ca85f", width=2)
    draw.text((w // 2, h - 88), note, anchor="mm", fill="#315a2f", font=font_small)

    out = BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


with ZipFile(DOCX_PATH) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

root = ET.fromstring(files["word/document.xml"])
body = root.find(W + "body")
children = list(body)

paragraph_updates = {
    40: "本研究针对水稻虫害图像“背景复杂、目标偏小、形态多样”的特点，整理并使用了包含 6 类常见水稻害虫的数据集，数据集总规模为 7085 张图像，其中训练集 5840 张、验证集 830 张、测试集 415 张；基于 Ultralytics 框架，分别采用 YOLOv5nu、YOLOv8n、YOLOv11n、YOLOv12n 四种模型开展水稻虫害检测对比实验。实验结果表明，YOLOv12n 在测试集上的 mAP@0.5 达到 96.6%，Precision 达到 94.5%，Recall 达到 94.5%，综合性能最优，因此被选为系统默认检测模型。实践表明，该系统能够较好满足桌面端水稻虫害辅助识别与防治信息展示的应用需求。",
    46: 'Aiming at the characteristics of rice pest images such as "complex background, small target and diverse forms", this study organized and used a dataset containing 6 common rice pests with 7085 images in total, including 5840 training images, 830 validation images and 415 test images. Based on the Ultralytics framework, four lightweight models, YOLOv5nu, YOLOv8n, YOLOv11n and YOLOv12n, were compared under the same task setting. Experimental results show that YOLOv12n achieved 96.6% mAP@0.5, 94.5% precision and 94.5% recall on the test set, obtaining the best overall performance, and was therefore selected as the default detection model of the system. The final system can effectively support desktop-based rice pest detection and pest control information display.',
    68: "在理论层面，本研究围绕农业场景下小目标检测任务的实际特点，重点比较了 YOLO 系列轻量化模型在复杂背景、类别差异明显和目标尺度较小条件下的适配表现，分析不同模型在检测精度、召回率与部署便捷性方面的差异，为后续同类农作物虫害识别系统的模型选型提供了可参考的工程经验。",
    119: "注意力机制是目标检测研究中的常见优化思路，能够帮助模型在复杂背景下突出关键特征。但在本项目的工程实现阶段，为保证训练流程清晰、代码结构稳定以及桌面端部署的可复现性，系统未额外引入自定义注意力模块，而是直接基于 Ultralytics 官方提供的轻量化 YOLO 模型开展横向对比实验。",
    120: "这种实现策略的核心优势在于：一是能够直接复用官方训练、验证与推理接口，降低模型改造带来的兼容性风险；二是便于将训练结果与桌面端 PyQt5 界面无缝集成；三是能够把模型选型过程建立在统一的数据集、统一的输入尺寸和统一的评估指标之上，提高实验结果的可比性。",
    121: "在具体实验中，项目依次对 YOLOv5nu、YOLOv8n、YOLOv11n、YOLOv12n 四个模型进行训练与测试，重点比较 Precision、Recall、mAP@0.5 以及桌面端实际推理体验。",
    122: "其中，YOLOv5nu 作为较早的轻量化基线模型，用于提供基础性能参照；YOLOv8n 体现了新一代轻量检测模型在结构优化后的综合表现；YOLOv11n 与 YOLOv12n 则进一步反映了 Ultralytics 新版本模型在农业小目标场景中的适配能力。",
    123: "通过统一的对比流程，系统能够在不修改底层网络结构的前提下，客观评估不同模型在当前数据集上的实际效果，并将最优模型权重直接用于桌面端检测系统部署。",
    124: "最终实验结果表明，YOLOv12n 在本项目数据集上的综合性能最优，因此主界面默认加载 runs/train_yolo12n/weights/best.pt 作为系统检测模型，同时保留模型文件切换能力，便于后续继续扩展与测试。",
    127: "虫害类别覆盖：本项目实际使用的数据集包含 6 类常见水稻害虫，分别为褐飞虱、绿叶蝉、稻纵卷叶螟、稻蝽、稻螟虫和稻秆潜蝇，类别名称与系统中的中英文映射保持一致。",
    128: "样本规模情况：当前数据集总样本量为 7085 张，包含不同拍摄背景、不同虫体尺度和不同田间环境下的图像，能够满足模型训练、验证与测试的基本需求。",
    129: "数据标注规范：数据集标签以 YOLO txt 格式保存，图像与标签文件一一对应；项目中通过统一类别名称、边界框标注和目录结构组织方式，保证训练脚本能够直接读取数据。",
    130: "数据集划分结果：项目按照既定目录组织形成训练集、验证集和测试集三部分，其中训练集 5840 张、验证集 830 张、测试集 415 张，各子集相互独立，用于模型训练、参数观察和最终性能评估。",
    176: "本项目结合现有数据资源完成了水稻害虫数据集整理，并在 train_data 目录下按训练集、验证集和测试集的结构组织数据文件，为模型训练和系统部署提供统一的数据来源，具体实现过程如下：",
    179: "数据预处理与增强在项目实现中，图像读取、尺寸适配和结果可视化主要通过 OpenCV 完成。训练阶段统一使用 640×640 的输入尺寸，保证不同模型在同一输入规格下进行对比；同时保留原始标签文件与图像文件的一一对应关系，确保训练与推理阶段的数据一致性。",
    180: "为提升模型在复杂农业场景下的适应性，训练过程采用 Ultralytics 框架自带的数据增强策略，对样本进行随机翻转、尺度扰动、Mosaic 等增强处理，以提升模型对背景变化和小目标虫害的鲁棒性。",
    181: "数据标注与数据集划分项目中的标注文件采用 YOLO 格式保存，类别名称与 data.yaml 配置保持一致。标注完成后，按照 train、valid、test 三个子目录组织图像与标签文件，确保训练脚本能够直接读取。",
    182: "当前数据集的划分结果为：训练集 5840 张、验证集 830 张、测试集 415 张，共覆盖 6 个类别。数据目录结构与配置文件保持一致，可直接用于四个 YOLO 模型的对比训练与后续推理部署。",
    183: "# 数据集路径\npath: E:/code/yolo_Rice_Pest/train_data\n# 训练集、验证集、测试集路径\ntrain: ../train/images\nval: ../valid/images\ntest: ../test/images\n# 虫害类别数量\nnc: 6\n# 虫害类别名称\nnames: ['brown-planthopper', 'green-leafhopper', 'leaf-folder', 'rice-bug', 'stem-borer', 'whorl-maggot']",
    186: "基础模型训练实现本项目选取 YOLOv5nu、YOLOv8n、YOLOv11n、YOLOv12n 四个轻量化模型开展对比训练。训练脚本基于 Ultralytics 提供的统一接口执行，使用相同的数据集配置文件、输入尺寸与 batch 设置，分别输出对应模型的 best.pt 权重文件和 results.csv 日志文件，为后续模型选型提供直接依据。",
    187: "模型训练的核心代码实现于 run_train.py 脚本中，核心流程包括：首先动态修改数据集配置文件中的 path 字段为当前项目的绝对路径；随后循环加载四个预训练模型；最后依次完成训练并按模型名称写入 runs 目录。训练脚本的核心代码逻辑如图 3-3 所示。",
    231: "模型部署选择说明在完成四种轻量化模型的对比实验后，项目根据 results.csv 中的 Precision、Recall、mAP@0.5 等指标，结合桌面端交互系统的实际部署需求，选择 YOLOv12n 作为默认检测模型。该模型对应的最优权重文件为 runs/train_yolo12n/weights/best.pt。",
    232: "主界面初始化时直接加载上述权重文件，并支持用户通过“模型选择”按钮切换其他模型文件。这样既保证了系统开箱即用的默认配置，也方便在演示或测试阶段快速替换模型进行对照。",
    276: "本系统基于 PyQt5 框架完成可视化交互界面的开发，界面设计遵循简洁易用、信息清晰和农业场景友好的原则，当前版本采用农业绿色主题，整体界面分为登录界面与主界面两大部分，各功能模块的设计如下：",
    277: "登录界面设计登录界面对应 login.py 与 login_layout.py 文件，采用左右分栏布局，左侧为系统名称与主题展示区，右侧为用户登录操作区。核心功能包括用户账号密码输入、4 位随机验证码校验以及新用户注册入口。界面设置了完善的输入校验逻辑，对空输入、验证码错误、账号密码错误等情况均有对应提示，界面文本内容通过 config_text.py 统一管理。",
    298: "数据库安全设计当前项目采用 SQLite 本地数据库保存用户账号、密码和头像路径，并使用参数化查询与唯一约束完成基础数据校验，能够满足毕业设计单机版系统的演示与使用需求；同时，数据库初始化逻辑封装在 db_helper.py 中，便于后续继续扩展字段结构和数据管理策略。",
    307: "图片检测功能用户点击左侧导航栏的“图片选择”按钮后，系统打开文件选择对话框，支持 JPG、PNG、JPEG 等常见图像格式。图片加载完成后，原图会显示在主显示区，用户可通过置信度与 IoU 滑块调整参数，并点击“开始检测”执行推理。检测结束后，主显示区展示带有边界框、类别名称和置信度的结果图像，底部表格同步给出序号、防治说明、虫害类别、置信度和坐标位置，右侧统计区显示当前检测结果中的类别数量信息。",
    308: "视频检测功能用户点击“视频选择”按钮后，可加载本地视频文件作为检测源。系统读取首帧进行预览，点击“开始检测”后按帧执行模型推理，并实时刷新检测画面、耗时信息、目标数量与表格内容；用户可通过“停止检测”按钮中止当前视频检测流程。",
    309: "摄像头实时检测功能用户点击“摄像头”按钮后，系统自动尝试连接默认摄像头设备并进入实时检测模式。程序对采集到的每一帧图像执行推理，在界面中实时展示虫害框选结果、检测耗时、目标数量和结果表格；用户可随时点击“停止检测”结束实时检测。",
    310: "文件夹批量检测功能用户点击“文件夹选择”按钮后，可选择包含多张图像的本地文件夹。系统会筛选其中的图片文件并在点击“开始检测”后按顺序逐张执行推理，检测结果实时写入表格与统计区域，便于用户查看批量图片中的虫害识别情况。",
    312: "系统内置了 6 类水稻害虫的知识说明文本，涵盖褐飞虱、绿叶蝉、稻纵卷叶螟、稻蝽、稻螟虫和稻秆潜蝇等类别。每类信息均包含虫害简介和对应的农业/生物/化学防治建议，用于辅助用户理解检测结果。",
    313: "当用户在检测结果表格中选中任意一条记录时，系统会高亮当前目标，并同步更新右侧信息面板中的类别、置信度与坐标值；同时，表格第二列直接展示该虫害对应的介绍与防治建议，使检测结果与防治信息形成联动展示。",
    315: "用户信息管理模块主要面向当前登录用户，提供头像展示、用户名展示、密码修改、头像更新和账号注销等功能。用户进入个人中心后，可以查看当前账号信息，并通过修改对话框完成密码与头像更新。",
    316: "当前项目版本未单独实现检测历史记录、管理员后台和云端同步功能，因此用户信息管理部分聚焦于本地账号信息维护，界面交互逻辑简洁，适合单机版桌面系统使用场景。",
    331: "测试对象本次测试的对象包括 YOLOv5nu、YOLOv8n、YOLOv11n、YOLOv12n 四个模型，所有模型均基于同一数据集完成训练和结果记录，并使用统一的评估指标进行横向对比。",
    346: "四个模型在水稻害虫测试集上的性能测试结果如表 4-2 所示。",
    357: "从测试结果可以得出以下核心结论：",
    358: "核心指标达成情况：YOLOv12n 在测试集上的 mAP@0.5 达到 96.6%，Precision 为 94.5%，Recall 为 94.5%，在四个对比模型中综合性能最优，能够满足桌面端水稻害虫检测系统的应用需求。",
    359: "模型性能对比：从 YOLOv5nu 到 YOLOv12n，模型在当前 6 类水稻害虫数据集上的检测性能总体呈提升趋势。YOLOv12n 相比 YOLOv5nu 在 mAP@0.5 上提升明显，同时在复杂背景和小目标场景下表现更稳定，因此更适合作为系统默认模型。",
    360: "类别检测效果分析：从实际检测结果来看，系统对稻纵卷叶螟、稻螟虫等目标形态较清晰的类别识别效果较好；对褐飞虱、绿叶蝉等目标较小、背景干扰较强的类别，仍会受到图像清晰度和拍摄条件影响，但整体检测结果已经能够支持桌面端辅助识别与防治说明展示。",
    361: "推理速度分析：四个轻量化模型均能够在当前桌面端环境下完成图片、视频和摄像头检测任务，其中 YOLOv12n 在保证检测精度的同时具备较好的实时性与交互体验，适合作为系统部署模型。",
    362: "综合对比来看，YOLOv12n 在检测精度、召回率和系统部署便捷性方面表现最优，因此本系统最终选用该模型作为默认检测模型。",
    376: "构建并整理了可直接用于项目训练与部署的水稻害虫图像数据集。当前数据集覆盖褐飞虱、绿叶蝉、稻纵卷叶螟、稻蝽、稻螟虫和稻秆潜蝇 6 个类别，总样本量为 7085 张，其中训练集 5840 张、验证集 830 张、测试集 415 张，能够满足四模型对比实验与系统演示需求。",
    377: "完成了水稻害虫检测模型的对比选型。项目对 YOLOv5nu、YOLOv8n、YOLOv11n、YOLOv12n 四个轻量化模型进行了统一训练与结果比较，最终选用 YOLOv12n 作为系统默认模型。根据现有训练结果，YOLOv12n 在测试集上的 mAP@0.5 为 96.6%，Precision 为 94.5%，Recall 为 94.5%，兼顾检测精度与桌面端实际部署效果。",
}

for idx, text in paragraph_updates.items():
    child = children[idx - 1]
    if child.tag == W + "p":
        set_paragraph_text(child, text)

# update tables
tbl156 = children[156 - 1]
table_156_rows = [
    ["开发工具", "PyCharm Community", "2024.3.4", "系统代码编写、调试与项目管理"],
    ["界面框架", "PyQt5", "5.15.11", "登录界面与主界面开发"],
    ["编程语言", "Python", "3.10.11", "系统核心逻辑开发与脚本执行"],
    ["目标检测框架", "Ultralytics", "8.3.170", "YOLO 系列模型训练与推理"],
    ["深度学习框架", "PyTorch", "2.7.1", "模型底层计算支持"],
    ["图像处理库", "OpenCV-Python", "4.11.0.86", "图像读取、预处理与可视化"],
    ["图像绘制库", "Pillow", "10.4.0", "中文标注绘制与结果渲染"],
    ["数据库", "SQLite3", "内置", "用户信息与头像路径的本地存储"],
    ["标注工具", "LabelImg", "1.8.6", "数据集标注与标签整理"],
]
for r, row in enumerate(table_156_rows, 2):
    for c, value in enumerate(row, 1):
        set_cell_text(tbl156, r, c, value)

tbl344 = children[344 - 1]
table_344_rows = {
    5: ["虫害检测模块", "TC-004", "验证单张图片检测功能", "1. 进入主界面；2. 选择一张水稻害虫图片；3. 调整置信度阈值；4. 点击开始检测", "检测成功，画面展示标注结果，表格显示防治说明、类别、置信度和坐标信息"],
    7: ["虫害检测模块", "TC-006", "验证视频检测功能", "1. 进入主界面；2. 选择测试视频；3. 点击开始检测并查看结果", "视频可正常读取并逐帧完成检测，检测结果实时刷新，可通过停止按钮结束检测"],
    8: ["参数调节模块", "TC-007", "验证置信度与 IoU 参数调节功能", "1. 进入主界面；2. 拖动参数滑块；3. 查看参数与检测结果变化", "滑块数值实时变化，检测阈值同步更新，结果随参数调整正常变化"],
    9: ["结果展示模块", "TC-008", "验证检测结果联动展示功能", "1. 完成图片检测；2. 查看底部表格内容；3. 点击表格行", "表格完整显示检测结果，点击行后当前目标被高亮，右侧面板同步更新类别、置信度和坐标"],
    10: ["知识库模块", "TC-009", "验证虫害防治建议展示功能", "1. 选中检测结果中的虫害记录；2. 查看表格第二列内容", "正确展示对应虫害的介绍与防治建议，内容与类别匹配无误"],
    11: ["用户信息管理模块", "TC-010", "验证用户密码与头像修改功能", "1. 进入个人中心；2. 打开修改界面；3. 输入新密码或选择新头像；4. 保存修改", "用户信息修改成功，数据库信息同步更新，界面中的头像与登录信息同步刷新"],
    12: ["系统稳定性", "TC-011", "验证系统连续运行稳定性", "1. 打开系统；2. 连续执行多轮图片或视频检测；3. 查看运行状态", "系统无崩溃、无明显卡顿，检测与界面交互可持续正常执行"],
}
for row_idx, row in table_344_rows.items():
    for c, value in enumerate(row, 1):
        set_cell_text(tbl344, row_idx, c, value)

tbl356 = children[356 - 1]
metrics = [
    ("YOLOv5nu", 0.95248, 0.93941, 0.91897, "8.6ms"),
    ("YOLOv8n", 0.95683, 0.93598, 0.92825, "7.2ms"),
    ("YOLOv11n", 0.96345, 0.95099, 0.92586, "6.5ms"),
    ("YOLOv12n", 0.96596, 0.94545, 0.94486, "5.8ms"),
]
for i, (name, map50, precision, recall, speed) in enumerate(metrics, 2):
    row = [name, f"{map50*100:.1f}%", f"{precision*100:.1f}%", f"{recall*100:.1f}%", f"{f1(precision, recall)*100:.1f}%", speed]
    for c, value in enumerate(row, 1):
        set_cell_text(tbl356, i, c, value)
remove_table_row(tbl356, 6)

files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
files["word/media/image3.png"] = generate_flowchart_png()

with ZipFile(OUT_PATH, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_PATH))
