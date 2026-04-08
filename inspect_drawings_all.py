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

for idx, node in enumerate(children, 1):
    if node.tag == W + "p" and node.findall(".//w:drawing", NS):
        text = get_text(node)
        media = []
        for blip in node.findall(".//a:blip", NS):
            rid = blip.attrib.get("{%s}embed" % NS["r"])
            if rid:
                media.append(rels.get(rid, rid))
        prev_text = get_text(children[idx - 2]) if idx >= 2 and children[idx - 2].tag == W + "p" else ""
        next_text = get_text(children[idx]) if idx < len(children) and children[idx].tag == W + "p" else ""
        print("drawing", idx, "text=", repr(text), "media=", media)
        print("  prev=", repr(prev_text))
        print("  next=", repr(next_text))
