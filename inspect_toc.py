from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    styles = ET.fromstring(zf.read("word/styles.xml"))

style_names = {}
for style in styles.findall("w:style", NS):
    style_id = style.attrib.get(W + "styleId")
    name = style.find("w:name", NS)
    style_names[style_id] = name.attrib.get(W + "val") if name is not None else style_id

body = doc.find("w:body", NS)
children = list(body)

print("== First 80 paragraphs ==")
for idx, child in enumerate(children[:80], 1):
    if child.tag != W + "p":
        continue
    text = get_text(child).strip()
    ppr = child.find("w:pPr", NS)
    style_id = "(none)"
    if ppr is not None:
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is not None:
            style_id = pstyle.attrib.get(W + "val")
    field_chars = [fc.attrib.get(W + "fldCharType") for fc in child.findall(".//w:fldChar", NS)]
    instr = "".join(t.text or "" for t in child.findall(".//w:instrText", NS))
    if text or field_chars or instr:
        print(idx, style_id, style_names.get(style_id, style_id), repr(text), field_chars, repr(instr))

print("\n== TOC styles ==")
for style in styles.findall("w:style", NS):
    style_id = style.attrib.get(W + "styleId")
    if style_id and style_id.upper().startswith("TOC"):
        name = style.find("w:name", NS)
        ppr = style.find("w:pPr", NS)
        rpr = style.find("w:rPr", NS)
        print("style", style_id, name.attrib.get(W + "val") if name is not None else "")
        if ppr is not None:
            print(" ppr", ET.tostring(ppr, encoding="unicode"))
        if rpr is not None:
            print(" rpr", ET.tostring(rpr, encoding="unicode"))
