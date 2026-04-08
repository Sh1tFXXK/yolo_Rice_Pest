from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    rels_xml = ET.fromstring(zf.read("word/_rels/document.xml.rels"))

rels = {rel.attrib.get("Id"): rel.attrib.get("Target") for rel in rels_xml}
body = doc.find("w:body", NS)
children = list(body)

targets = [168, 169, 170, 171, 172, 173, 174, 228, 229, 230, 231, 232, 272, 273, 274, 275, 286, 287, 288, 289, 340, 341, 342, 343, 352, 353, 354, 355, 388, 389, 390, 391, 392, 393, 394, 395]
for idx in targets:
    if idx > len(children):
        continue
    node = children[idx - 1]
    kind = "tbl" if node.tag == W + "tbl" else "p" if node.tag == W + "p" else node.tag
    text = get_text(node)
    media = []
    for blip in node.findall(".//a:blip", NS):
        rid = blip.attrib.get("{%s}embed" % NS["r"])
        if rid:
            media.append(rels.get(rid, rid))
    print(idx, kind, repr(text[:120]), media)
