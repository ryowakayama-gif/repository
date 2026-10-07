# -*- coding: utf-8 -*-
"""docx に埋め込まれたネイティブグラフ（chartN.xml）の系列データを抜き出す。

使い方: python3 dump_docx_charts.py <入力docx> <出力txt>

本文中の登場順にグラフを並べ、直前の見出し（キャプション）と対応づける。
表の数値とグラフの数値が食い違っていないかの照合に使う。
"""
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)


def chart_order(path):
    """本文順の (キャプション, chart部品名) の並びを返す"""
    doc = Document(path)
    with zipfile.ZipFile(path) as z:
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    rid2part = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))

    order = []
    cap = ""
    for blk in iter_blocks(doc):
        if not isinstance(blk, Paragraph):
            continue
        t = blk.text.strip()
        ids = re.findall(r'r:id="(rId\d+)"',
                         blk._p.xml.replace(R, "r:"))
        ids += re.findall(r'{[^}]*relationships}id="(rId\d+)"', blk._p.xml)
        ids = [i for i in ids if rid2part.get(i, "").startswith("charts/")]
        if ids:
            for i in ids:
                order.append((cap, "word/" + rid2part[i]))
        elif t:
            cap = t
    return order


def read_chart(z, part):
    x = ET.fromstring(z.read(part))
    tt = x.find(f"{C}chart/{C}title")
    title = "".join(
        e.text or "" for e in
        tt.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")
    ) if tt is not None else ""
    cats, out = [], []
    for ser in x.iter(C + "ser"):
        nm = ser.find(f"{C}tx//{C}v")
        name = nm.text if nm is not None else ""
        cs = [v.text for v in ser.findall(f"{C}cat//{C}pt/{C}v")]
        vs = [v.text for v in ser.findall(f"{C}val//{C}pt/{C}v")]
        if cs and not cats:
            cats = cs
        out.append((name, vs))
    return title, cats, out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    order = chart_order(src)
    lines = [f"元ファイル：{src}", f"グラフ数：{len(order)}", ""]
    with zipfile.ZipFile(src) as z:
        for i, (cap, part) in enumerate(order, 1):
            try:
                title, cats, sers = read_chart(z, part)
            except KeyError:
                lines.append(f"[グラフ{i}] {part} を読めません")
                continue
            lines.append(f"■ グラフ{i}（{part}）　直前の見出し：{cap}")
            if title:
                lines.append(f"　タイトル：{title}")
            for name, vs in sers:
                pairs = []
                for j, v in enumerate(vs):
                    c = cats[j] if j < len(cats) else f"#{j}"
                    pairs.append(f"{c}={v}")
                lines.append(f"　系列「{name}」：" + "、".join(pairs))
            lines.append("")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"{dst}: グラフ{len(order)}点")


if __name__ == "__main__":
    main()
