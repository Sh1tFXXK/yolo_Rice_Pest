from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from copy import deepcopy

SRC_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_working.docx")
IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v2.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v3.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def get_children(docx_path):
    with ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    return root, body, list(body)


src_root, src_body, src_children = get_children(SRC_DOCX)
src_drawing_para = deepcopy(src_children[172 - 1])

with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

root = ET.fromstring(files["word/document.xml"])
body = root.find("w:body", NS)
children = list(body)

# Replace duplicated caption paragraph with the original drawing paragraph.
body.remove(children[173 - 1])
body.insert(173 - 1, src_drawing_para)

files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
