from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    styles = ET.fromstring(zf.read("word/styles.xml"))
    settings = ET.fromstring(zf.read("word/settings.xml"))

print("== updateFields ==")
node = settings.find("w:updateFields", NS)
print(None if node is None else node.attrib.get(W + "val"))

print("\n== TOC styles ==")
for sid in ["12", "13", "9"]:
    for style in styles.findall("w:style", NS):
        if style.attrib.get(W + "styleId") == sid:
            name = style.find("w:name", NS)
            ppr = style.find("w:pPr", NS)
            rpr = style.find("w:rPr", NS)
            print("style", sid, None if name is None else name.attrib.get(W + "val"))
            print(" ppr", ET.tostring(ppr, encoding="unicode") if ppr is not None else None)
            print(" rpr", ET.tostring(rpr, encoding="unicode") if rpr is not None else None)

body = doc.find("w:body", NS)
children = list(body)

print("\n== Sections with page format ==")
for idx, sect in enumerate(doc.findall(".//w:sectPr", NS), 1):
    pg = sect.find("w:pgNumType", NS)
    sec_type = sect.find("w:type", NS)
    print(idx, "fmt=", None if pg is None else pg.attrib.get(W + "fmt"), "type=", None if sec_type is None else sec_type.attrib.get(W + "val"))

print("\n== TOC first paragraphs ==")
for child in body.findall("w:sdt", NS):
    gallery = child.find(".//w:docPartGallery", NS)
    if gallery is None or gallery.attrib.get(W + "val") != "Table of Contents":
        continue
    content = child.find("w:sdtContent", NS)
    for idx, p in enumerate(content.findall("w:p", NS)[:6], 1):
        text = get_text(p)
        ppr = p.find("w:pPr", NS)
        tabs = ppr.find("w:tabs", NS) if ppr is not None else None
        spacing = ppr.find("w:spacing", NS) if ppr is not None else None
        ind = ppr.find("w:ind", NS) if ppr is not None else None
        jc = ppr.find("w:jc", NS) if ppr is not None else None
        print(idx, repr(text))
        print(" spacing", None if spacing is None else spacing.attrib)
        print(" ind", None if ind is None else ind.attrib)
        print(" jc", None if jc is None else jc.attrib)
        if tabs is not None:
            for tab in tabs.findall("w:tab", NS):
                print(" tab", tab.attrib)
    break

print("\n== Reference and acknowledgement styles ==")
for idx in range(388, 411):
    if idx > len(children):
        break
    node = children[idx - 1]
    if node.tag != W + "p":
        continue
    text = get_text(node)
    if text in {"参考文献", "致　谢", "致谢"}:
        ppr = node.find("w:pPr", NS)
        pstyle = None if ppr is None else ppr.find("w:pStyle", NS)
        print(idx, text, None if pstyle is None else pstyle.attrib.get(W + "val"))
