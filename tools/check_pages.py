# -*- coding: utf-8 -*-
"""docx の紙面を実際に数えて、体裁の崩れを見つける。

使い方
    python3 tools/check_pages.py <pdf> <docx>

PDF は docx を LibreOffice で変換したもの。

    soffice --headless --convert-to pdf --outdir <dir> <docx>

数えるもの
  - 総ページ数
  - 中身の少ないページ（表紙・最終ページを除き400字未満）
  - 紙面をまたぐ表（表の先頭の行と最後の行が別のページにあるもの）

表の位置は、表題（「表N　…」）があればそれで、なければ
文書の並び順に沿って先頭の行・最後の行の文字を探して定める。
"""

import re
import sys

import pymupdf
from docx import Document


def norm(s):
    return re.sub(r"[\s　]", "", s or "")


def main(pdf, docx):
    pages = [norm(p.get_text()) for p in pymupdf.open(pdf)]

    def pageof(s, frm=1):
        s = norm(s)
        if not s:
            return None
        for i in range(max(frm, 1) - 1, len(pages)):
            if s in pages[i]:
                return i + 1
        return None

    doc = Document(docx)
    capmap = {}
    for p in doc.paragraphs:
        m = re.match(r"表(\d+)[　 ]", p.text)
        if m:
            capmap[int(m.group(1))] = p.text

    split, cur = [], 1
    for n, t in enumerate(doc.tables, 1):
        # 表の先頭を探す手掛かり。表題があればそれを使う。
        head = capmap.get(n) or (
            "".join(c.text for c in t.rows[0].cells)
            + "".join(c.text for c in t.rows[1].cells)
            if len(t.rows) > 1 else "")
        last = "".join(c.text for c in t.rows[-1].cells)[:60]
        pa = pageof(head, cur)
        pb = pageof(last, pa or cur)
        if pa:
            cur = pa
        if pa and pb and pa != pb:
            split.append((n, pa, pb, (capmap.get(n) or last)[:28]))

    short = [(i, len(t)) for i, t in enumerate(pages, 1)
             if len(t) < 400 and 1 < i < len(pages)]
    print("総ページ %d" % len(pages))
    print("中身の少ないページ（表紙・最終ページを除く）:", short)
    print("紙面をまたぐ表 %d件" % len(split))
    for x in split:
        print("   表%d 先頭p%d→最終行p%d  %s" % x)
    return len(split)


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1], sys.argv[2]) else 0)
