from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def text_from_node(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX_PATH) as zf:
    xml = zf.read("word/document.xml")
    media = [n for n in zf.namelist() if n.startswith("word/media/")]

root = ET.fromstring(xml)
body = root.find("w:body", NS)
children = list(body)

for idx in [40, 46, 119, 127, 183, 231, 276, 298, 312, 315, 331, 358, 377]:
    node = children[idx - 1]
    print(f"P{idx}: {text_from_node(node)}")

for t_idx in [156, 344, 356]:
    tbl = children[t_idx - 1]
    print(f"T{t_idx}:")
    for r, tr in enumerate(tbl.findall(W + "tr"), 1):
        cells = [text_from_node(tc) for tc in tr.findall(W + "tc")]
        print(f"  R{r}: {' | '.join(cells)}")

print("MEDIA:", media)
