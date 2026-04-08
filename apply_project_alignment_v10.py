from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET


IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v9.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v10.docx")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def ensure(parent, tag, first=False):
    node = parent.find(tag, NS)
    if node is None:
        node = ET.Element(W + tag.split(":")[1])
        if first and len(parent):
            parent.insert(0, node)
        else:
            parent.append(node)
    return node


def first_run_rpr(p):
    run = p.find("w:r", NS)
    if run is None:
        return None
    rpr = run.find("w:rPr", NS)
    return deepcopy(rpr) if rpr is not None else None


def clear_except_ppr(p):
    for child in list(p):
        if child.tag != W + "pPr":
            p.remove(child)


def set_paragraph_text(p, text):
    ppr = p.find("w:pPr", NS)
    rpr = first_run_rpr(p)
    clear_except_ppr(p)
    if text:
        run = ET.Element(W + "r")
        if rpr is not None:
            run.append(rpr)
        t = ET.SubElement(run, W + "t")
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        p.append(run)
    if ppr is not None and (not len(p) or p[0] is not ppr):
        if ppr in p:
            p.remove(ppr)
        p.insert(0, ppr)


def set_cell_text(tbl, row_idx, col_idx, text):
    tr = tbl.findall("w:tr", NS)[row_idx - 1]
    tc = tr.findall("w:tc", NS)[col_idx - 1]
    paragraphs = tc.findall("w:p", NS)
    if not paragraphs:
        paragraphs = [ET.SubElement(tc, W + "p")]
    set_paragraph_text(paragraphs[0], text)
    for extra in paragraphs[1:]:
        tc.remove(extra)


def remove_table_rows_after(tbl, keep_rows):
    rows = tbl.findall("w:tr", NS)
    for row in rows[keep_rows:]:
        tbl.remove(row)


PARAGRAPH_UPDATES = {
    40: "本研究围绕桌面端水稻害虫识别需求，设计并实现了一套基于 YOLOv12n 的水稻害虫检测系统。项目实际使用的数据集共包含 6 类常见水稻害虫，图像总量为 7085 张，其中训练集 5840 张、验证集 830 张、测试集 415 张。系统采用 PyQt5 构建可视化界面，提供用户注册登录、图片检测、视频检测、摄像头实时检测、文件夹批量检测、结果表格联动展示以及虫害防治建议查询等功能。结合项目训练日志中的最终验证结果，YOLOv12n 的 Precision 为 94.5%，Recall 为 94.5%，mAP@0.5 为 96.6%，能够满足当前桌面端水稻害虫辅助识别与信息展示的应用需求。",
    46: "This study focuses on the desktop application scenario of rice pest recognition and implements a rice pest detection system based on YOLOv12n. The project dataset contains 6 common rice pest categories and 7085 images in total, including 5840 training images, 830 validation images and 415 test images. The system uses PyQt5 to build a visual interface and provides user registration and login, image detection, video detection, real-time camera detection, folder batch detection, linked result display and pest control suggestion query. According to the final validation record in the training log, YOLOv12n achieves 94.5% precision, 94.5% recall and 96.6% mAP@0.5, which is sufficient for the current desktop-based rice pest assisted identification task.",
    59: "水稻虫害检测是智慧农业领域的重要研究方向之一，其核心任务是通过计算机视觉与深度学习技术，自动识别水稻图像中的虫害类别与位置，替代传统人工巡查中的重复性识别工作，为田间虫害诊断和防治提供辅助依据。本章节围绕水稻虫害检测的研究背景、研究意义与国内外研究现状展开系统阐述，明确本研究的核心目标与技术路线。",
    68: "在理论层面，本研究结合农业场景中“小目标、复杂背景、类别差异明显”的任务特点，对桌面端虫害识别系统中目标检测模型的工程落地方式进行了梳理，并围绕项目实际采用的 YOLOv12n 模型分析其在水稻害虫识别场景中的适配性，为同类农业视觉系统的模型选型与部署提供参考。",
    69: "在应用层面，本研究开发的可视化水稻害虫检测系统，可直接服务于普通农户、农业种植合作社和基层农技推广人员。系统支持图片、视频、摄像头实时、文件夹批量四种检测模式，无需用户具备专业的深度学习背景，通过简单的界面操作即可完成虫害识别，并同步查看类别统计、目标详情以及对应的防治建议，具有较强的实用价值。",
    107: "2.3 本项目采用的检测模型",
    108: "2.3.1 YOLOv12n 模型",
    109: "YOLO（You Only Look Once）系列模型是一类典型的一阶段目标检测算法，能够在一次前向传播中同时完成目标定位与类别识别，具有结构清晰、部署便捷和实时性较好的特点。结合本项目桌面端应用场景，需要兼顾检测精度、推理速度与集成难度，因此论文内容仅围绕项目实际部署所使用的 YOLOv12n 模型展开说明。",
    111: "YOLOv12n 是 Ultralytics 框架提供的轻量化检测模型版本，适合在普通 PC 环境下完成虫害图像的快速识别。该模型能够在保持较低计算开销的同时完成多尺度目标检测，较适合水稻害虫这种目标较小、背景较复杂的图像场景。",
    114: "在本项目中，YOLOv12n 的优势主要体现在三个方面：一是能够直接通过 Ultralytics 官方接口完成训练、验证与推理，减少额外改模带来的兼容性问题；二是模型权重文件便于在 PyQt5 主界面中直接加载，适合桌面端集成；三是训练日志与结果文件结构清晰，便于后续分析和维护。",
    117: "结合当前项目实现，系统默认加载 runs/train_yolo12n/weights/best.pt 作为检测模型，并通过 model.predict 接口完成图片、视频、摄像头和文件夹图像的统一推理。检测结果再经过类别名称映射与界面联动更新，形成完整的桌面端识别流程。",
    118: "2.3.2 模型部署说明",
    119: "本项目在模型部署阶段未额外引入自定义注意力模块或改写网络结构，而是直接采用 Ultralytics 官方提供的 YOLOv12n 模型进行训练与集成。这样做的原因在于项目重点是完成一个可运行、可演示、易维护的桌面端虫害检测系统，而不是进行复杂的网络结构创新。",
    120: "这种实现方式能够直接复用官方训练、验证与推理接口，降低模型改造带来的适配风险，也便于将训练结果与 PyQt5 主界面无缝集成。对于毕业设计场景而言，这种工程实现路线更贴近项目实际，也更容易保证系统的稳定性与可复现性。",
    121: "在程序启动后，主界面会默认加载 YOLOv12n 最优权重文件，并根据当前媒体类型执行统一的推理流程。用户在界面中可以通过“模型选择”按钮手动切换其他权重文件，但系统默认说明始终围绕已实际使用的 YOLOv12n 模型展开。",
    122: "模型推理时，系统会读取当前图片或视频帧数据，根据界面滑块提供的置信度阈值和 IoU 阈值执行预测，再将结果解析为类别、置信度和边界框坐标等结构化信息，用于后续可视化展示。",
    123: "随后，程序将英文类别名称映射为中文虫害名称，联动更新结果图像、底部表格、类别统计列表以及右侧目标详情面板，使模型输出能够直接服务于用户查看和分析。",
    124: "因此，YOLOv12n 在本项目中的作用不仅是完成底层识别任务，同时也是连接训练结果、界面交互和农业知识展示的核心算法模块。",
    168: "数据存储模块：基于 SQLite 数据库实现，负责存储用户账号信息与头像路径等本地数据，满足单机版桌面系统的基础数据管理需求。",
    169: "系统整体采用模块化组织方式，当前已经实现默认模型加载、手动模型切换、虫害知识内容内置展示等功能，代码结构清晰，便于后续在真实需求基础上继续扩展。",
    176: "本项目结合现有数据资源完成了水稻害虫数据集整理，并在 train_data 目录下按训练集、验证集和测试集组织图像与标签文件，为 YOLOv12n 模型训练和系统部署提供统一的数据来源，具体实现过程如下：",
    179: "数据预处理与增强在项目实现中，图像读取、尺寸适配和结果可视化主要通过 OpenCV 完成。训练阶段统一使用 640×640 的输入尺寸，保证模型输入规格一致；同时保留图像文件与标签文件的一一对应关系，确保训练与推理阶段的数据结构保持统一。",
    180: "为提升模型对复杂农业场景的适应性，训练过程采用 Ultralytics 框架自带的数据增强策略，对样本进行随机翻转、尺度扰动、Mosaic 等增强处理，以提高模型对背景变化和小目标虫害的鲁棒性。",
    181: "数据标注与数据集划分项目中的标注文件采用 YOLO 格式保存，类别名称与 data.yaml 配置保持一致。标注完成后，按照 train、valid、test 三个子目录组织图像与标签文件，保证训练脚本和后续推理流程能够直接读取。",
    182: "当前数据集的划分结果为：训练集 5840 张、验证集 830 张、测试集 415 张，共覆盖 6 个类别。该目录结构与配置文件保持一致，可直接用于 YOLOv12n 模型训练、验证和部署。",
    185: "本研究基于 Ultralytics 框架完成检测算法的实现，论文内容围绕系统当前实际部署所使用的 YOLOv12n 模型展开，核心涉及数据配置、模型训练、权重加载以及界面推理集成四个环节，具体实现过程如下：",
    186: "模型训练实现项目使用 Ultralytics 提供的统一接口加载预训练权重并执行训练流程。结合当前项目目录，系统默认部署所对应的训练结果保存在 runs/train_yolo12n 目录下，其中包括最优权重文件 best.pt、训练日志 results.csv 以及训练参数文件 args.yaml。",
    187: "训练相关逻辑由 run_train.py 脚本负责组织，核心过程包括更新 data.yaml 中的数据路径、加载预训练权重、执行模型训练以及输出训练结果文件。虽然工程目录中保留了多个模型文件，论文只围绕系统实际使用的 YOLOv12n 模型说明其部署过程。训练脚本的核心代码逻辑如图 3-3 所示。",
    231: "模型部署说明当前桌面端系统默认使用的检测模型为 YOLOv12n，对应权重文件路径为 runs/train_yolo12n/weights/best.pt。该权重在主界面初始化阶段直接加载，用于后续图片、视频、摄像头和文件夹批量检测。",
    232: "同时，系统保留了“模型选择”按钮，允许用户在演示或调试阶段手动切换其他 pt 权重文件。但论文中的算法说明与性能分析仅围绕项目实际默认部署的 YOLOv12n 模型展开，不再延伸到其他模型说明。",
    233: "推理检测功能统一集成在主界面的 MainWindow 类中，系统提供单张图片、视频文件、摄像头实时和文件夹批量四种检测模式，用户加载媒体后点击“开始检测”即可进入对应的推理流程。",
    234: "推理检测的核心流程为：首先加载默认或用户选择的模型权重文件；其次读取当前图像、视频帧或摄像头画面；随后调用 predict 接口执行前向推理，获得检测结果；最后将结果解析为类别、置信度和坐标信息，并联动更新图像标注、结果表格、类别统计和右侧详情面板。",
    281: "核心检测区：位于界面中央，是系统的核心交互区域，分为三个子模块：顶部为参数调节区，包含置信度与 IoU 阈值的滑动调节控件；中部为检测画面展示区，通过 QLabel 展示原始图像、视频帧或检测后的标注结果；底部为检测结果表格区，采用 QTableWidget 展示序号、害虫介绍及防治、虫害类别、置信度、坐标位置五列信息，用户可点击表格行查看对应目标详情。",
    286: "本系统采用轻量级的 SQLite 数据库完成用户数据的本地存储与管理，无需额外安装数据库服务，适配单机版部署需求。数据库相关操作全部封装于 db_helper.py 中，实际实现内容包括数据库初始化、用户名存在检查、用户注册写入、登录校验以及用户信息更新等功能。",
    316: "当前项目版本未单独实现检测历史记录、Excel 导出、管理员后台、GIS 地图展示和云端同步等扩展功能，因此用户信息管理部分聚焦于本地账号信息维护，整体设计更贴合当前单机版桌面系统的实际实现状态。",
    328: "4.2.1 YOLOv12n 检测效果测试",
    329: "检测效果测试的核心目标是说明当前项目默认模型 YOLOv12n 在现有数据集和训练配置下的实际表现，并判断该模型是否能够满足桌面端虫害识别系统的部署需求。本文以 runs/train_yolo12n 目录中的训练结果文件为依据，结合界面中的实际推理流程进行分析。",
    330: "测试指标本节采用 Precision、Recall、mAP@0.5 和 mAP@0.5:0.95 四项指标作为模型效果说明依据。这些指标均来自 Ultralytics 训练日志中的验证集记录，能够直接反映当前默认模型的检测能力。",
    331: "测试对象本节测试对象为系统实际默认加载的 YOLOv12n 模型，对应权重文件为 runs/train_yolo12n/weights/best.pt，训练参数记录于 runs/train_yolo12n/args.yaml，结果记录于 runs/train_yolo12n/results.csv。",
    332: "测试流程检测效果测试的具体流程如下：",
    333: "读取 runs/train_yolo12n/args.yaml，确认当前训练参数为 epochs=100、imgsz=640、batch=32；",
    334: "读取 runs/train_yolo12n/results.csv 的最终一轮验证记录，提取 Precision、Recall、mAP@0.5 和 mAP@0.5:0.95 指标；",
    335: "结合主界面中的图片、视频、摄像头和文件夹检测流程，检查模型权重是否能够被正常加载并完成推理展示；",
    336: "汇总关键指标并形成表 4-2，用于说明系统默认模型的检测效果；",
    337: "根据训练日志记录与界面运行效果，评估当前模型是否满足桌面端部署需求。",
    346: "",
    352: "4.3.1 YOLOv12n 检测效果分析",
    353: "YOLOv12n 模型在当前项目训练日志中的最终验证结果如表 4-2 所示。",
    354: "表4-2 YOLOv12n 模型验证结果",
    357: "结合 runs/train_yolo12n/results.csv 中最后一轮训练记录，可以得到以下结论：",
    358: "在第 100 轮训练结束时，YOLOv12n 的 Precision 为 94.5%，Recall 为 94.5%，mAP@0.5 为 96.6%，mAP@0.5:0.95 为 68.4%，说明当前模型已经具备较好的目标检测能力。",
    359: "从项目实际识别任务来看，YOLOv12n 能够完成 6 类水稻害虫的目标定位与类别识别，适合当前桌面端系统的默认部署需求。",
    360: "对于褐飞虱、绿叶蝉等目标较小且背景干扰较强的类别，检测效果仍会受到图像清晰度、遮挡情况和拍摄角度影响，但整体结果已经能够支持虫害辅助识别与防治信息展示。",
    361: "结合主界面中的图片、视频、摄像头和文件夹检测流程，当前默认模型在实际运行中能够完成统一推理与结果联动展示，满足单机版桌面系统的核心使用需求。",
    362: "因此，系统在部署阶段直接加载 YOLOv12n 权重文件，并保留模型文件手动切换入口作为后续调试与扩展手段。",
    364: "结合表 4-1 所列的 11 项核心功能测试，可以看出当前项目版本已经覆盖登录注册、四种检测模式、参数调节、结果联动展示、虫害知识说明以及用户信息维护等主要功能。",
    365: "在用户认证与信息管理方面，系统已经实现验证码校验、用户注册、登录验证、密码修改、头像更新和账号注销等完整流程，相关数据通过 SQLite 本地数据库进行管理。",
    366: "在检测功能方面，图片、视频、摄像头和文件夹四种模式均在主界面中形成统一的操作逻辑：先加载媒体，再点击“开始检测”执行推理，用户能够清晰地完成整个识别过程。",
    367: "在结果展示方面，检测耗时、目标数量、类别统计、结果表格和右侧详情面板能够联动更新，表格第二列同步展示对应虫害的介绍与防治建议，使识别结果具备更直接的应用价值。",
    368: "在参数与交互方面，置信度和 IoU 滑块能够实时更新检测阈值，类别过滤和表格行选中功能可以正常驱动界面刷新，提升了结果查看的便捷性。",
    369: "需要说明的是，当前项目并未实现管理员后台、历史记录、Excel 导出和云端同步等扩展功能，因此论文中的功能分析仅围绕已经完成的桌面端能力展开。",
    370: "总体而言，当前项目已经完成毕业设计所需的核心桌面端功能，能够满足单机版水稻害虫检测、结果展示与防治建议查询的基本需求。",
    371: "后续若继续迭代，可在现有实现基础上进一步优化不同分辨率下的界面布局表现与模型运行效率。",
    376: "构建并整理了可直接用于项目训练与部署的水稻害虫图像数据集。当前数据集覆盖褐飞虱、绿叶蝉、稻纵卷叶螟、稻蝽、稻螟虫和稻秆潜蝇 6 个类别，总样本量为 7085 张，其中训练集 5840 张、验证集 830 张、测试集 415 张，能够满足当前模型训练与桌面系统演示需求。",
    377: "完成了基于 YOLOv12n 的检测模型训练与部署。项目默认加载 runs/train_yolo12n/weights/best.pt 作为系统检测模型，结合训练日志中的最终验证结果，Precision 为 94.5%，Recall 为 94.5%，mAP@0.5 为 96.6%，说明该模型能够满足当前桌面端虫害识别任务的基本要求。",
    379: "围绕项目当前版本设计了 11 项核心功能测试用例，覆盖登录注册、四种检测模式、参数调节、结果展示、知识说明和用户信息管理等模块。结合代码实现和界面交互流程分析，可以看出系统核心功能完整，运行逻辑清晰，具备较好的实用性。",
    383: "数据集的进一步扩充与优化：未来可以继续补充更多田间实拍样本，丰富不同光照、遮挡和复杂背景下的虫害图像，进一步提升模型在真实农业场景中的泛化能力。",
    384: "模型与推理流程优化：后续可在保持当前工程结构稳定的前提下，继续优化模型训练参数与推理效率，提升系统在普通 PC 环境下处理视频和摄像头画面的流畅性。",
    385: "界面交互细节优化：未来可以围绕不同分辨率屏幕下的布局适配、提示信息展示和结果查看方式继续完善界面体验，使桌面端操作更加直观。",
    386: "工程化部署完善：后续可进一步整理运行环境、打包启动流程和模型文件管理方式，降低系统部署门槛，提升项目在教学展示和实际演示场景中的可用性。",
}

BLANK_PARAGRAPHS = {135, 155, 158, 289, 343, 355}

TABLE_156_ROWS = [
    ["类别", "工具 / 框架", "版本号", "用途说明"],
    ["开发工具", "PyCharm Community", "2024.3.4", "系统代码编写、调试与项目管理"],
    ["界面框架", "PyQt5", "5.15.11", "登录界面与主界面开发"],
    ["编程语言", "Python", "3.10.11", "系统核心逻辑开发与脚本执行"],
    ["目标检测框架", "Ultralytics", "8.3.170", "YOLOv12n 模型训练与推理"],
    ["深度学习框架", "PyTorch", "2.7.1", "模型底层计算支持"],
    ["图像处理库", "OpenCV-Python", "4.11.0.86", "图像读取、预处理与可视化"],
    ["图像绘制库", "Pillow", "10.4.0", "中文标注绘制与结果渲染"],
    ["数据库", "SQLite3", "内置", "用户信息与头像路径的本地存储"],
    ["配置解析", "PyYAML", "6.0.2", "数据集配置读取与写回"],
]

TABLE_356_ROWS = [
    ["模型", "Precision", "Recall", "mAP@0.5", "mAP@0.5:0.95", "训练参数"],
    ["YOLOv12n", "94.5%", "94.5%", "96.6%", "68.4%", "epochs=100；imgsz=640；batch=32"],
]


with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

doc = ET.fromstring(files["word/document.xml"])
body = doc.find("w:body", NS)
children = list(body)

for idx, text in PARAGRAPH_UPDATES.items():
    node = children[idx - 1]
    if node.tag == W + "p":
        set_paragraph_text(node, text)

for idx in BLANK_PARAGRAPHS:
    node = children[idx - 1]
    if node.tag == W + "p" and get_text(node) == "表格":
        set_paragraph_text(node, "")

tbl156 = children[156 - 1]
for r, row in enumerate(TABLE_156_ROWS, 1):
    for c, value in enumerate(row, 1):
        set_cell_text(tbl156, r, c, value)

tbl356 = children[356 - 1]
remove_table_rows_after(tbl356, 2)
for r, row in enumerate(TABLE_356_ROWS, 1):
    for c, value in enumerate(row, 1):
        set_cell_text(tbl356, r, c, value)

files["word/document.xml"] = ET.tostring(doc, encoding="utf-8", xml_declaration=True)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
