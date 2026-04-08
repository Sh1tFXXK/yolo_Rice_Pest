from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    for name in zf.namelist():
        if name.startswith("word/header") or name.startswith("word/footer"):
            root = ET.fromstring(zf.read(name))
            print("==", name, "==")
            for idx, p in enumerate(root.findall("w:p", NS), 1):
                print(idx, repr(get_text(p)))
                print(ET.tostring(p, encoding="unicode"))
