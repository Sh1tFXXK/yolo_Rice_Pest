from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
import re


IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]

TOC_STYLES = {
    "12": 0,   # toc 1
    "13": 200, # toc 2
    "9": 400,  # toc 3
}
TOC_LEFT_TWIPS = {
    0: 0,
    200: 420,
    400: 840,
}


def ensure(parent, tag, first=False):
    node = parent.find(tag, NS)
    if node is None:
        node = ET.Element(W + tag.split(":")[1])
        if first and len(parent):
            parent.insert(0, node)
        else:
            parent.append(node)
    return node


def remove_all(parent, tag):
    for node in list(parent.findall(tag, NS)):
        parent.remove(node)


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def set_run_font(run, font="宋体", size=24):
    rpr = ensure(run, "w:rPr", first=True)
    rfonts = ensure(rpr, "w:rFonts")
    rfonts.set(W + "ascii", font)
    rfonts.set(W + "hAnsi", font)
    rfonts.set(W + "eastAsia", font)
    rfonts.set(W + "cs", font)
    sz = ensure(rpr, "w:sz")
    sz.set(W + "val", str(size))
    szcs = ensure(rpr, "w:szCs")
    szcs.set(W + "val", str(size))
    color = ensure(rpr, "w:color")
    color.set(W + "val", "000000")


def set_spacing(ppr, line="500", before="0", after="0"):
    spacing = ensure(ppr, "w:spacing")
    spacing.set(W + "line", line)
    spacing.set(W + "lineRule", "exact")
    spacing.set(W + "before", before)
    spacing.set(W + "after", after)


def set_indent(ppr, left_chars=None, first_line_chars=0):
    ind = ensure(ppr, "w:ind")
    if left_chars is None:
        ind.attrib.pop(W + "leftChars", None)
        ind.attrib.pop(W + "left", None)
    else:
        ind.set(W + "leftChars", str(left_chars))
        ind.set(W + "left", str(TOC_LEFT_TWIPS.get(int(left_chars), 0)))
    ind.set(W + "firstLineChars", str(first_line_chars))
    ind.set(W + "firstLine", "0")


def set_jc(ppr, val):
    jc = ensure(ppr, "w:jc")
    jc.set(W + "val", val)


def set_tab_right_dot(ppr, pos="8844"):
    remove_all(ppr, "w:tabs")
    tabs = ET.SubElement(ppr, W + "tabs")
    tab = ET.SubElement(tabs, W + "tab")
    tab.set(W + "val", "right")
    tab.set(W + "leader", "dot")
    tab.set(W + "pos", pos)


def set_paragraph_style(p, style_id):
    ppr = ensure(p, "w:pPr", first=True)
    pstyle = ensure(ppr, "w:pStyle")
    pstyle.set(W + "val", style_id)


def normalize_page_result(text):
    stripped = text.strip()
    m = re.fullmatch(r"-\s*(\d+)\s*-", stripped)
    if m:
        return m.group(1)
    return text


with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

doc = ET.fromstring(files["word/document.xml"])
styles = ET.fromstring(files["word/styles.xml"])
settings = ET.fromstring(files["word/settings.xml"])


# 1. TOC styles in styles.xml
for style in styles.findall("w:style", NS):
    style_id = style.attrib.get(W + "styleId")
    if style_id not in TOC_STYLES:
        continue
    ppr = ensure(style, "w:pPr")
    set_spacing(ppr, line="500", before="0", after="0")
    set_indent(ppr, left_chars=TOC_STYLES[style_id], first_line_chars=0)
    set_tab_right_dot(ppr, pos="8844")
    rpr = ensure(style, "w:rPr")
    rfonts = ensure(rpr, "w:rFonts")
    rfonts.set(W + "ascii", "宋体")
    rfonts.set(W + "hAnsi", "宋体")
    rfonts.set(W + "eastAsia", "宋体")
    rfonts.set(W + "cs", "宋体")
    sz = ensure(rpr, "w:sz")
    sz.set(W + "val", "24")
    szcs = ensure(rpr, "w:szCs")
    szcs.set(W + "val", "24")


# 2. Make fields update on open
update_fields = settings.find("w:updateFields", NS)
if update_fields is None:
    update_fields = ET.SubElement(settings, W + "updateFields")
update_fields.set(W + "val", "true")


body = doc.find("w:body", NS)
children = list(body)


# 3. Fix all section page numbering to plain Arabic digits
for sect in doc.findall(".//w:sectPr", NS):
    pg_num = sect.find("w:pgNumType", NS)
    if pg_num is None:
        pg_num = ET.SubElement(sect, W + "pgNumType")
    pg_num.set(W + "fmt", "decimal")


# 4. Ensure "参考文献" and "致谢" remain in auto TOC
for child in children:
    if child.tag != W + "p":
        continue
    text = get_text(child)
    if text in {"参考文献", "致　谢", "致谢"}:
        set_paragraph_style(child, "2")


# 5. Fix TOC content control
toc_sdt = None
for sdt in body.findall("w:sdt", NS):
    gallery = sdt.find(".//w:docPartGallery", NS)
    if gallery is not None and gallery.attrib.get(W + "val") == "Table of Contents":
        toc_sdt = sdt
        break

if toc_sdt is not None:
    sdt_pr = toc_sdt.find("w:sdtPr", NS)
    if sdt_pr is not None:
        rpr = ensure(sdt_pr, "w:rPr")
        rfonts = ensure(rpr, "w:rFonts")
        rfonts.set(W + "ascii", "宋体")
        rfonts.set(W + "hAnsi", "宋体")
        rfonts.set(W + "eastAsia", "宋体")
        rfonts.set(W + "cs", "宋体")
        sz = ensure(rpr, "w:sz")
        sz.set(W + "val", "24")
        szcs = ensure(rpr, "w:szCs")
        szcs.set(W + "val", "24")

    content = toc_sdt.find("w:sdtContent", NS)
    if content is not None:
        toc_children = list(content)
        for p in toc_children:
            if p.tag != W + "p":
                continue
            text = get_text(p)
            ppr = ensure(p, "w:pPr", first=True)

            if text == "目录":
                set_spacing(ppr, line="500", before="0", after="0")
                set_indent(ppr, left_chars=0, first_line_chars=0)
                set_jc(ppr, "center")
                remove_all(ppr, "w:tabs")
                for run in p.findall("w:r", NS):
                    set_run_font(run, font="宋体", size=24)
                continue

            style_id = None
            pstyle = ppr.find("w:pStyle", NS)
            if pstyle is not None:
                style_id = pstyle.attrib.get(W + "val")

            if style_id in TOC_STYLES:
                set_spacing(ppr, line="500", before="0", after="0")
                set_indent(ppr, left_chars=TOC_STYLES[style_id], first_line_chars=0)
                set_jc(ppr, "left")
                set_tab_right_dot(ppr, pos="8844")

                for run in p.findall("w:r", NS):
                    set_run_font(run, font="宋体", size=24)
                    for t in run.findall("w:t", NS):
                        t.text = normalize_page_result(t.text or "")


# 6. Make the section after TOC start on an odd page
for idx, child in enumerate(children):
    if child is toc_sdt and idx + 1 < len(children):
        next_node = children[idx + 1]
        if next_node.tag == W + "p":
            ppr = ensure(next_node, "w:pPr", first=True)
            sect = ensure(ppr, "w:sectPr")
            sec_type = sect.find("w:type", NS)
            if sec_type is None:
                sec_type = ET.SubElement(sect, W + "type")
            sec_type.set(W + "val", "oddPage")
        break


files["word/document.xml"] = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
files["word/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
files["word/settings.xml"] = ET.tostring(settings, encoding="utf-8", xml_declaration=True)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
