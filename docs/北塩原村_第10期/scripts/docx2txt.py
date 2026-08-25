# -*- coding: utf-8 -*-
"""docx → プレーンテキスト（表はセル区切りで出力）"""
import sys, zipfile, re
from xml.etree import ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def para_text(p):
    out = []
    for node in p.iter():
        if node.tag == W+"t":
            out.append(node.text or "")
        elif node.tag == W+"tab":
            out.append("\t")
        elif node.tag == W+"br":
            out.append("\n")
    return "".join(out)

def walk(el, depth=0):
    lines = []
    for child in el:
        if child.tag == W+"p":
            t = para_text(child).strip()
            if t: lines.append(t)
        elif child.tag == W+"tbl":
            lines.append("<<TABLE>>")
            for tr in child.findall(W+"tr"):
                cells = []
                for tc in tr.findall(W+"tc"):
                    ct = " ".join(para_text(p).strip() for p in tc.findall(W+"p"))
                    cells.append(re.sub(r"\s+", " ", ct).strip())
                lines.append(" | ".join(cells))
            lines.append("<<END TABLE>>")
        elif child.tag in (W+"sdt", W+"body", W+"sdtContent", W+"txbxContent"):
            lines.extend(walk(child, depth))
    return lines

path = sys.argv[1]
with zipfile.ZipFile(path) as z:
    root = ET.fromstring(z.read("word/document.xml"))
body = root.find(W+"body")
print("\n".join(walk(body)))
