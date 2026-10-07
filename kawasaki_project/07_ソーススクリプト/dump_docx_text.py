# -*- coding: utf-8 -*-
"""docx を本文＋表のテキストに落とす（レビュー用）。

使い方: python3 dump_docx_text.py <入力docx> <出力txt>

本文段落と表を文書順に出力する。表はパイプ区切りで、表番号を付ける。
グラフ（inline shape / chart part）がある段落には [図] マークを付ける。
"""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


def iter_block_items(parent):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)


def cell_text(c):
    return " ".join(p.text.strip() for p in c.paragraphs if p.text.strip())


def main():
    src, dst = sys.argv[1], sys.argv[2]
    doc = Document(src)
    out = []
    ti = 0
    pi = 0
    for blk in iter_block_items(doc):
        if isinstance(blk, Paragraph):
            pi += 1
            t = blk.text.strip()
            has_draw = blk._p.findall('.//' + qn('w:drawing'))
            if t:
                sty = blk.style.name if blk.style is not None else ""
                mark = "[図]" if has_draw else ""
                out.append(f"[P{pi}]{mark}<{sty}> {t}")
            elif has_draw:
                out.append(f"[P{pi}][図（キャプションなし）]")
        else:
            ti += 1
            out.append(f"--- TABLE {ti} ({len(blk.rows)}行 x {len(blk.columns)}列) ---")
            for r in blk.rows:
                cells = []
                seen = set()
                for c in r.cells:
                    key = id(c._tc)
                    if key in seen:
                        continue
                    seen.add(key)
                    cells.append(cell_text(c))
                out.append(" | ".join(cells))
            out.append(f"--- /TABLE {ti} ---")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print(f"{dst}: 段落{pi} 表{ti} 行{len(out)}")


if __name__ == "__main__":
    main()
