from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified.docx")
OUT_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v2.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
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


def make_flowchart():
    width, height = 1100, 760
    img = Image.new("RGB", (width, height), "#f5f8f1")
    draw = ImageDraw.Draw(img)

    try:
        ft_title = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 30)
        ft_stage = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 17)
        ft_box = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 19)
        ft_small = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 14)
    except Exception:
        ft_title = ft_stage = ft_box = ft_small = ImageFont.load_default()

    def box(x, y, w, h, text, fill, stroke="#2E6031", text_fill="white", radius=18, font=ft_box):
        draw.rounded_rectangle((x, y, x + w, y + h), radius=radius, fill=fill, outline=stroke, width=3)
        draw.multiline_text((x + w / 2, y + h / 2), text, anchor="mm", align="center", fill=text_fill, font=font, spacing=6)

    def stage(x, y, w, text):
        box(x, y, w, 42, text, "#DDEFD2", stroke="#7CA85F", text_fill="#315A2F", radius=14, font=ft_stage)

    def arrow(points, fill="#3E7A3E", width=5):
        draw.line(points, fill=fill, width=width)
        ex, ey = points[-1]
        if len(points) >= 2:
            sx, sy = points[-2]
            if abs(ex - sx) > abs(ey - sy):
                if ex > sx:
                    tri = [(ex, ey), (ex - 14, ey - 8), (ex - 14, ey + 8)]
                else:
                    tri = [(ex, ey), (ex + 14, ey - 8), (ex + 14, ey + 8)]
            else:
                if ey > sy:
                    tri = [(ex, ey), (ex - 8, ey - 14), (ex + 8, ey - 14)]
                else:
                    tri = [(ex, ey), (ex - 8, ey + 14), (ex + 8, ey + 14)]
            draw.polygon(tri, fill=fill)

    draw.rounded_rectangle((18, 18, width - 18, height - 18), radius=24, outline="#6F9C55", width=3, fill="#eef5e8")
    draw.text((width / 2, 48), "图 3-1  水稻害虫检测系统业务流程图", anchor="mm", fill="#2E6031", font=ft_title)

    stage(55, 95, 235, "阶段一  用户进入系统")
    box(85, 155, 170, 58, "步骤1\n打开登录界面", "#4C8F49")
    box(60, 255, 220, 72, "步骤2\n输入用户名、密码和验证码", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F")
    box(88, 372, 165, 86, "步骤3\n校验登录信息", "#FFF8E8", stroke="#9A8B52", text_fill="#5A4D1F")
    box(70, 505, 115, 56, "校验失败\n返回登录页", "#F6E7CF", stroke="#B88A4A", text_fill="#6C4A1F", radius=14, font=ft_small)
    box(200, 505, 115, 56, "校验成功\n进入主界面", "#5F9F52", text_fill="white", radius=14, font=ft_small)

    stage(345, 95, 245, "阶段二  选择检测任务")
    box(380, 155, 175, 64, "步骤4\n加载默认模型\nYOLOv12n", "#4C8F49")
    box(388, 255, 160, 58, "步骤5\n选择检测源", "#6EA85A")
    box(305, 360, 120, 54, "图片检测", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F", radius=14, font=ft_small)
    box(455, 360, 120, 54, "视频检测", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F", radius=14, font=ft_small)
    box(305, 440, 120, 60, "文件夹\n批量检测", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F", radius=14, font=ft_small)
    box(455, 440, 120, 60, "摄像头\n实时检测", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F", radius=14, font=ft_small)

    stage(670, 95, 290, "阶段三  执行推理与结果联动")
    box(695, 155, 240, 68, "步骤6\n设置置信度、IoU\n并可切换模型文件", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F")
    box(732, 265, 166, 56, "步骤7\n执行模型推理", "#5F9F52")
    box(692, 365, 246, 60, "步骤8\n绘制检测框、类别和置信度", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F")
    box(692, 465, 246, 72, "步骤9\n更新结果表格\n序号 / 防治说明 / 类别 / 置信度 / 坐标", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F")
    box(692, 580, 246, 72, "步骤10\n联动更新右侧面板\n类别统计 / 当前目标详情", "#F7FBF3", stroke="#7CA85F", text_fill="#315A2F")
    box(735, 686, 160, 48, "完成一次检测流程", "#EAF4E2", stroke="#6F9C55", text_fill="#2E6031", radius=14, font=ft_small)

    arrow([(170, 213), (170, 255)])
    arrow([(170, 327), (170, 372)])
    arrow([(252, 415), (320, 415), (320, 533), (315, 533)])
    arrow([(146, 458), (146, 505)], fill="#A66B36")
    arrow([(226, 458), (258, 458), (258, 505)], fill="#3E7A3E")

    arrow([(315, 533), (350, 533), (350, 187), (380, 187)])
    arrow([(468, 219), (468, 255)])
    arrow([(468, 313), (365, 360)])
    arrow([(468, 313), (515, 360)])
    arrow([(468, 313), (365, 440)])
    arrow([(468, 313), (515, 440)])

    arrow([(425, 387), (650, 387), (650, 189), (695, 189)])
    arrow([(575, 387), (650, 387), (650, 189), (695, 189)])
    arrow([(425, 470), (650, 470), (650, 189), (695, 189)])
    arrow([(575, 470), (650, 470), (650, 189), (695, 189)])

    arrow([(815, 223), (815, 265)])
    arrow([(815, 321), (815, 365)])
    arrow([(815, 425), (815, 465)])
    arrow([(815, 537), (815, 580)])
    arrow([(815, 652), (815, 686)])

    note = "说明：系统支持图片、视频、文件夹和摄像头四种检测入口，所有入口最终汇入统一的 YOLO 推理与结果展示流程。"
    draw.rounded_rectangle((120, 620, 580, 710), radius=18, fill="#dceccd", outline="#7CA85F", width=2)
    draw.multiline_text((350, 665), note, anchor="mm", align="center", fill="#315A2F", font=ft_small, spacing=6)

    bio = BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


with ZipFile(DOCX_PATH) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

root = ET.fromstring(files["word/document.xml"])
body = root.find(W + "body")
children = list(body)

set_paragraph_text(children[171 - 1], "本系统的完整业务流程如图 3-1 所示。为了便于说明系统执行逻辑，流程图将项目运行过程拆解为十个连续步骤：步骤1为打开登录界面；步骤2为输入用户名、密码和验证码；步骤3为完成登录校验；步骤4为进入主界面并加载默认模型 YOLOv12n；步骤5为选择检测源；步骤6为设置置信度、IoU 或切换模型文件；步骤7为执行模型推理；步骤8为绘制检测框、类别与置信度；步骤9为更新结果表格；步骤10为联动更新右侧统计与目标详情信息，最终完成一次完整的检测流程。")
set_paragraph_text(children[172 - 1], "图 3-1 水稻害虫检测系统业务流程图")

files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
files["word/media/image3.png"] = make_flowchart()

with ZipFile(OUT_PATH, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_PATH))
