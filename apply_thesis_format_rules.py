from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
import re


IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = "{%s}" % NS["w"]

HEADING_STYLE_IDS = {"2", "3", "4", "5"}

STANDARD_REFERENCES = [
    "[1] 王健, 李明, 张华. 深度学习在水稻虫害识别中的应用研究[J]. 农业工程学报, 2024, 40(5): 12-20.",
    "[2] 刘畅. 基于YOLOv8的作物虫害识别系统设计与实现[D]. 北京: 中国农业大学, 2023.",
    "[3] 张晓峰, 王丽丽. 注意力机制优化的轻量化虫害识别模型研究[J]. 计算机应用研究, 2023, 40(8): 2345-2351.",
    "[4] 陈亮, 赵强. 基于深度学习的农业小目标虫害检测算法[J]. 计算机工程与设计, 2022, 43(11): 3456-3461.",
    "[5] 李娜, 吴涛. 多源数据融合的水稻虫害识别模型优化[J]. 信息技术, 2025, 39(2): 78-83.",
    "[6] 张海波. 智慧农业中虫害识别与预警系统的应用[J]. 农业机械学报, 2023, 54(7): 135-139.",
    "[7] 王小明, 陈晓华. 深度学习图像识别实战[M]. 北京: 电子工业出版社, 2022.",
]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def ensure(parent, tag, first=False):
    child = parent.find(tag, NS)
    if child is None:
        child = ET.Element(W + tag.split(":")[1])
        if first and len(parent):
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def get_ppr(p):
    return ensure(p, "w:pPr", first=True)


def get_rpr(run):
    return ensure(run, "w:rPr", first=True)


def clear_children(parent, keep_tags):
    for child in list(parent):
        if child.tag not in keep_tags:
            parent.remove(child)


def remove_child(parent, tag):
    for child in list(parent.findall(tag, NS)):
        parent.remove(child)


def set_rfonts(rpr, east_asia, ascii_font="Times New Roman", hansi_font="Times New Roman"):
    rfonts = ensure(rpr, "w:rFonts")
    rfonts.set(W + "ascii", ascii_font)
    rfonts.set(W + "hAnsi", hansi_font)
    rfonts.set(W + "eastAsia", east_asia)


def set_size(rpr, half_points):
    sz = ensure(rpr, "w:sz")
    sz.set(W + "val", str(half_points))
    szcs = ensure(rpr, "w:szCs")
    szcs.set(W + "val", str(half_points))


def set_bold(rpr, enabled):
    for tag in ("w:b", "w:bCs"):
        child = ensure(rpr, tag)
        child.set(W + "val", "1" if enabled else "0")


def set_run_font(run, east_asia, half_points, bold=None):
    rpr = get_rpr(run)
    set_rfonts(rpr, east_asia)
    set_size(rpr, half_points)
    color = ensure(rpr, "w:color")
    color.set(W + "val", "000000")
    if bold is not None:
        set_bold(rpr, bold)


def set_paragraph_text(p, text, east_asia="宋体", half_points=24, bold=None):
    ppr = p.find("w:pPr", NS)
    clear_children(p, {W + "pPr"})
    run = ET.Element(W + "r")
    t = ET.SubElement(run, W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    p.append(run)
    set_run_font(run, east_asia=east_asia, half_points=half_points, bold=bold)
    if ppr is not None and p[0] is not ppr:
        p.remove(ppr)
        p.insert(0, ppr)


def set_paragraph_format(
    p,
    *,
    align=None,
    first_line_chars=None,
    left_chars=None,
    hanging_chars=None,
    line=None,
    line_rule="exact",
    before=None,
    after=None,
    keep_next=False,
    keep_lines=False,
):
    ppr = get_ppr(p)
    if align is not None:
        jc = ensure(ppr, "w:jc")
        jc.set(W + "val", align)
    else:
        remove_child(ppr, "w:jc")

    ind = ppr.find("w:ind", NS)
    if first_line_chars is not None or left_chars is not None or hanging_chars is not None:
        if ind is None:
            ind = ET.SubElement(ppr, W + "ind")
        for attr, value in (
            ("firstLineChars", first_line_chars),
            ("leftChars", left_chars),
            ("hangingChars", hanging_chars),
        ):
            if value is None:
                ind.attrib.pop(W + attr, None)
            else:
                ind.set(W + attr, str(value))
    elif ind is not None:
        ppr.remove(ind)

    spacing = ppr.find("w:spacing", NS)
    if line is not None or before is not None or after is not None:
        if spacing is None:
            spacing = ET.SubElement(ppr, W + "spacing")
        if line is not None:
            spacing.set(W + "line", str(line))
            spacing.set(W + "lineRule", line_rule)
        else:
            spacing.attrib.pop(W + "line", None)
            spacing.attrib.pop(W + "lineRule", None)
        for attr, value in (("before", before), ("after", after)):
            if value is None:
                spacing.attrib.pop(W + attr, None)
            else:
                spacing.set(W + attr, str(value))
    elif spacing is not None:
        ppr.remove(spacing)

    for tag, enabled in (("w:keepNext", keep_next), ("w:keepLines", keep_lines)):
        node = ppr.find(tag, NS)
        if enabled:
            if node is None:
                ET.SubElement(ppr, W + tag.split(":")[1])
        elif node is not None:
            ppr.remove(node)


def normalize_caption(text):
    match = re.match(r"^(图|表)\s*(\d+-\d+)\s*(.*)$", text.strip())
    if not match:
        return text.strip().rstrip("。．.；;：:")
    kind, number, title = match.groups()
    title = title.strip().rstrip("。．.；;：:")
    return f"{kind}{number} {title}"


def replace_decimal_lists(text):
    return re.sub(r"(^|[；;])\s*(\d+)\.\s*", lambda m: f"{m.group(1)}（{m.group(2)}）", text)


def paragraph_style_id(p):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.attrib.get(W + "val")


def previous_nonempty_paragraph(children, idx):
    for pos in range(idx - 1, -1, -1):
        node = children[pos]
        if node.tag == W + "p":
            text = get_text(node)
            if text:
                return pos, node
    return None, None


def style_heading_paragraph(p):
    set_paragraph_format(
        p,
        align="left",
        first_line_chars=200,
        line=500,
        before=0,
        after=0,
    )
    for run in p.findall("w:r", NS):
        set_run_font(run, east_asia="黑体", half_points=24, bold=False)


def style_caption_paragraph(p):
    set_paragraph_format(
        p,
        align="center",
        first_line_chars=0,
        line=300,
        before=240,
        after=120,
        keep_lines=True,
    )
    for run in p.findall("w:r", NS):
        set_run_font(run, east_asia="宋体", half_points=21, bold=False)


def style_reference_heading(p):
    set_paragraph_format(
        p,
        align="left",
        first_line_chars=200,
        line=500,
        before=0,
        after=0,
    )
    for run in p.findall("w:r", NS):
        set_run_font(run, east_asia="黑体", half_points=24, bold=False)


def style_reference_entry(p):
    set_paragraph_format(
        p,
        align="left",
        first_line_chars=None,
        left_chars=200,
        hanging_chars=200,
        line=500,
        before=0,
        after=0,
    )
    for run in p.findall("w:r", NS):
        set_run_font(run, east_asia="宋体", half_points=24, bold=False)


def ensure_tc_borders(tc_pr, *, top=None, bottom=None, left=None, right=None):
    borders = tc_pr.find("w:tcBorders", NS)
    if borders is None:
        borders = ET.SubElement(tc_pr, W + "tcBorders")
    values = {"top": top, "bottom": bottom, "left": left, "right": right}
    for side, val in values.items():
        if val is None:
            continue
        node = borders.find(f"w:{side}", NS)
        if node is None:
            node = ET.SubElement(borders, W + side)
        if val == "nil":
            node.set(W + "val", "nil")
        else:
            node.set(W + "val", "single")
            node.set(W + "sz", "8")
            node.set(W + "color", "000000")
            node.set(W + "space", "0")


def style_table(tbl):
    tbl_pr = ensure(tbl, "w:tblPr", first=True)
    borders = tbl_pr.find("w:tblBorders", NS)
    if borders is None:
        borders = ET.SubElement(tbl_pr, W + "tblBorders")
    border_settings = {
        "top": "single",
        "bottom": "single",
        "left": "nil",
        "right": "nil",
        "insideH": "nil",
        "insideV": "nil",
    }
    for side, val in border_settings.items():
        node = borders.find(f"w:{side}", NS)
        if node is None:
            node = ET.SubElement(borders, W + side)
        node.set(W + "val", val)
        if val != "nil":
            node.set(W + "sz", "8")
            node.set(W + "color", "000000")
            node.set(W + "space", "0")

    rows = tbl.findall("w:tr", NS)
    for row_idx, row in enumerate(rows):
        cells = row.findall("w:tc", NS)
        for col_idx, cell in enumerate(cells):
            tc_pr = ensure(cell, "w:tcPr", first=True)
            ensure_tc_borders(tc_pr, top="nil", bottom="nil", left="nil", right="nil")
            if row_idx == 0:
                ensure_tc_borders(tc_pr, bottom="single")

            paragraphs = cell.findall("w:p", NS)
            for para in paragraphs:
                text = get_text(para)
                if text:
                    text = replace_decimal_lists(text)
                    set_paragraph_text(para, text, east_asia="宋体", half_points=21, bold=(row_idx == 0))
                set_paragraph_format(
                    para,
                    align="left" if col_idx == 0 else "center",
                    first_line_chars=0,
                    line=300,
                    before=0,
                    after=0,
                )
                for run in para.findall("w:r", NS):
                    set_run_font(run, east_asia="宋体", half_points=21, bold=(row_idx == 0))


with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

doc = ET.fromstring(files["word/document.xml"])
body = doc.find("w:body", NS)

# Move figure captions below their image paragraphs.
children = list(body)
idx = 0
while idx < len(children) - 1:
    current = children[idx]
    nxt = children[idx + 1]
    if current.tag == W + "p" and nxt.tag == W + "p":
        current_text = get_text(current)
        if current_text.startswith("图") and nxt.findall(".//w:drawing", NS):
            body.remove(nxt)
            body.remove(current)
            body.insert(idx, nxt)
            body.insert(idx + 1, current)
            children = list(body)
            idx += 2
            continue
    idx += 1

children = list(body)

# Headings and captions.
for child in children:
    if child.tag != W + "p":
        continue
    text = get_text(child)
    style_id = paragraph_style_id(child)

    if style_id in HEADING_STYLE_IDS:
        style_heading_paragraph(child)

    if text == "参考文献":
        style_reference_heading(child)

    if text.startswith("图") or text.startswith("表"):
        normalized = normalize_caption(text)
        set_paragraph_text(child, normalized, east_asia="宋体", half_points=21, bold=False)
        style_caption_paragraph(child)

    if child.findall(".//w:drawing", NS):
        set_paragraph_format(
            child,
            align="center",
            first_line_chars=0,
            line=300,
            before=240,
            after=0,
            keep_next=True,
            keep_lines=True,
        )

# Content tables: the previous non-empty paragraph is a table caption.
children = list(body)
for idx, child in enumerate(children):
    if child.tag != W + "tbl":
        continue
    prev_idx, prev_p = previous_nonempty_paragraph(children, idx)
    if prev_p is None:
        continue
    prev_text = get_text(prev_p)
    if prev_text.startswith("表"):
        style_table(child)

# Standardize reference entries.
children = list(body)
ref_start = None
for idx, child in enumerate(children):
    if child.tag == W + "p" and get_text(child) == "参考文献":
        ref_start = idx + 1
        break

if ref_start is not None:
    ref_paragraphs = []
    for child in children[ref_start:]:
        if child.tag == W + "p" and get_text(child):
            ref_paragraphs.append(child)
    for para, ref_text in zip(ref_paragraphs, STANDARD_REFERENCES):
        set_paragraph_text(para, ref_text, east_asia="宋体", half_points=24, bold=False)
        style_reference_entry(para)

files["word/document.xml"] = ET.tostring(doc, encoding="utf-8", xml_declaration=True)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
