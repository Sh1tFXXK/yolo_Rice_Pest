from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image, ImageDraw, ImageFont


IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v3.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
PNG_PATH = Path(r"E:\code\yolo_Rice_Pest\flowchart_formal_v6.png")


def load_font(size: int):
    for font_path in (
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ):
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            continue
    return ImageFont.load_default()


FT_TITLE = load_font(28)
FT_LABEL = load_font(19)
FT_TEXT = load_font(18)
FT_SMALL = load_font(16)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font, max_width: int):
    lines = []
    for paragraph in text.split("\n"):
        current = ""
        for ch in paragraph:
            candidate = current + ch
            if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = ch
        lines.append(current)
    return "\n".join(lines)


def draw_box(draw, xy, text, *, step=None, fill="#FFFFFF", outline="#5E8B5B", radius=12, font=FT_TEXT):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=2)

    if step:
        tag = (x1 + 12, y1 - 16, x1 + 82, y1 + 16)
        draw.rounded_rectangle(tag, radius=10, fill="#DCEED6", outline=outline, width=2)
        draw.text(((tag[0] + tag[2]) / 2, (tag[1] + tag[3]) / 2), step, font=FT_SMALL, fill="#2F5D31", anchor="mm")

    wrapped = wrap_text(draw, text, font, max_width=int(x2 - x1 - 26))
    draw.multiline_text(((x1 + x2) / 2, (y1 + y2) / 2), wrapped, font=font, fill="#1F1F1F", anchor="mm", align="center", spacing=4)


def draw_diamond(draw, center, width, height, text, *, step=None):
    cx, cy = center
    points = [(cx, cy - height // 2), (cx + width // 2, cy), (cx, cy + height // 2), (cx - width // 2, cy)]
    draw.polygon(points, fill="#FFFBEA", outline="#8C7A43")
    draw.line(points + [points[0]], fill="#8C7A43", width=2)
    if step:
        tag = (cx - 34, cy - height // 2 - 28, cx + 34, cy - height // 2 + 4)
        draw.rounded_rectangle(tag, radius=10, fill="#F1E8C4", outline="#8C7A43", width=2)
        draw.text(((tag[0] + tag[2]) / 2, (tag[1] + tag[3]) / 2), step, font=FT_SMALL, fill="#5D4E1D", anchor="mm")
    wrapped = wrap_text(draw, text, FT_SMALL, max_width=width - 40)
    draw.multiline_text((cx, cy), wrapped, font=FT_SMALL, fill="#3D3214", anchor="mm", align="center", spacing=3)


def draw_arrow(draw, points, *, fill="#466B48", width=3, label=None, label_offset=(0, 0)):
    draw.line(points, fill=fill, width=width)
    x1, y1 = points[-2]
    x2, y2 = points[-1]
    if abs(x2 - x1) >= abs(y2 - y1):
        if x2 >= x1:
            arrow = [(x2, y2), (x2 - 12, y2 - 7), (x2 - 12, y2 + 7)]
        else:
            arrow = [(x2, y2), (x2 + 12, y2 - 7), (x2 + 12, y2 + 7)]
    else:
        if y2 >= y1:
            arrow = [(x2, y2), (x2 - 7, y2 - 12), (x2 + 7, y2 - 12)]
        else:
            arrow = [(x2, y2), (x2 - 7, y2 + 12), (x2 + 7, y2 + 12)]
    draw.polygon(arrow, fill=fill)

    if label:
        lx = (points[-2][0] + points[-1][0]) / 2 + label_offset[0]
        ly = (points[-2][1] + points[-1][1]) / 2 + label_offset[1]
        box = (lx - 20, ly - 12, lx + 20, ly + 12)
        draw.rounded_rectangle(box, radius=8, fill="#FFFFFF", outline=fill, width=1)
        draw.text((lx, ly), label, font=FT_SMALL, fill=fill, anchor="mm")


def make_flowchart_png():
    img = Image.new("RGB", (1260, 930), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    draw.rectangle((24, 24, 1236, 906), outline="#93B58D", width=2)
    draw.text((630, 62), "图 3-1 水稻害虫检测系统业务流程图", font=FT_TITLE, fill="#274A2A", anchor="mm")

    draw.rounded_rectangle((60, 108, 360, 860), radius=18, outline="#C7DCC1", width=2, fill="#FAFCF8")
    draw.rounded_rectangle((390, 108, 720, 860), radius=18, outline="#C7DCC1", width=2, fill="#FAFCF8")
    draw.rounded_rectangle((750, 108, 1185, 860), radius=18, outline="#C7DCC1", width=2, fill="#FAFCF8")

    draw.text((210, 135), "阶段一  登录与进入系统", font=FT_LABEL, fill="#3D663E", anchor="mm")
    draw.text((555, 135), "阶段二  选择检测任务", font=FT_LABEL, fill="#3D663E", anchor="mm")
    draw.text((968, 135), "阶段三  模型推理与结果展示", font=FT_LABEL, fill="#3D663E", anchor="mm")

    draw_box(draw, (115, 180, 305, 242), "打开登录界面", step="步骤1")
    draw_box(draw, (90, 286, 330, 360), "输入用户名、密码和验证码", step="步骤2")
    draw_box(draw, (105, 404, 315, 466), "校验登录信息", step="步骤3")
    draw_diamond(draw, (210, 564), 190, 118, "登录是否成功")
    draw_box(draw, (95, 674, 205, 734), "返回登录页", fill="#FFF9ED", outline="#8C7A43")

    draw_box(draw, (435, 180, 675, 264), "进入主界面并加载默认模型\nYOLOv12n", step="步骤4")
    draw_box(draw, (465, 328, 645, 390), "选择检测源", step="步骤5")
    draw_box(draw, (470, 448, 640, 502), "图片检测", fill="#F9FCF7")
    draw_box(draw, (470, 530, 640, 584), "视频检测", fill="#F9FCF7")
    draw_box(draw, (470, 612, 640, 666), "文件夹检测", fill="#F9FCF7")
    draw_box(draw, (470, 694, 640, 748), "摄像头检测", fill="#F9FCF7")

    draw_box(draw, (835, 182, 1100, 266), "设置置信度、IoU\n或切换模型文件", step="步骤6")
    draw_box(draw, (870, 326, 1065, 388), "执行模型推理", step="步骤7")
    draw_box(draw, (825, 446, 1110, 518), "绘制检测框、类别和置信度", step="步骤8")
    draw_box(draw, (825, 568, 1110, 650), "更新结果表格\n序号 / 防治说明 / 类别 / 置信度 / 坐标", step="步骤9")
    draw_box(draw, (825, 708, 1110, 790), "更新右侧统计与目标详情\n类别统计 / 当前目标信息", step="步骤10")

    draw_arrow(draw, [(210, 242), (210, 286)])
    draw_arrow(draw, [(210, 360), (210, 404)])
    draw_arrow(draw, [(210, 466), (210, 505)])
    draw_arrow(draw, [(115, 564), (95, 564), (95, 704)], label="失败", label_offset=(0, -18), fill="#8C7A43")
    draw_arrow(draw, [(305, 564), (380, 564), (380, 222), (435, 222)], label="成功", label_offset=(0, -18))

    draw_arrow(draw, [(555, 264), (555, 328)])
    draw_arrow(draw, [(555, 390), (555, 448)])
    draw_arrow(draw, [(555, 502), (555, 530)])
    draw_arrow(draw, [(555, 584), (555, 612)])
    draw_arrow(draw, [(555, 666), (555, 694)])

    draw.line([(700, 455), (700, 740)], fill="#466B48", width=3)
    draw_arrow(draw, [(640, 475), (700, 475)])
    draw_arrow(draw, [(640, 557), (700, 557)])
    draw_arrow(draw, [(640, 639), (700, 639)])
    draw_arrow(draw, [(640, 721), (700, 721)])
    draw_arrow(draw, [(700, 598), (790, 598), (790, 224), (835, 224)])

    draw_arrow(draw, [(968, 266), (968, 326)])
    draw_arrow(draw, [(968, 388), (968, 446)])
    draw_arrow(draw, [(968, 518), (968, 568)])
    draw_arrow(draw, [(968, 650), (968, 708)])

    draw.rounded_rectangle((850, 826, 1086, 874), radius=12, fill="#F7FBF4", outline="#5E8B5B", width=2)
    draw.text((968, 850), "完成一次完整检测流程", font=FT_SMALL, fill="#2F5D31", anchor="mm")
    draw_arrow(draw, [(968, 790), (968, 826)])

    bio = BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

png_bytes = make_flowchart_png()
files["word/media/image3.png"] = png_bytes
PNG_PATH.write_bytes(png_bytes)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
