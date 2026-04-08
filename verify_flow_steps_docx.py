from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v2.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def text_from_node(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX_PATH) as zf:
    xml = zf.read("word/document.xml")
    info = zf.getinfo("word/media/image3.png")

root = ET.fromstring(xml)
body = root.find("w:body", NS)
children = list(body)

print("P171:", text_from_node(children[171 - 1]))
print("P172:", text_from_node(children[172 - 1]))
print("image3.png bytes:", info.file_size)
