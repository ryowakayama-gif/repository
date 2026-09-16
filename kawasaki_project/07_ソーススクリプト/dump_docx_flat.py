# -*- coding: utf-8 -*-
"""1行1レコードを保証する dump（セル内改行は ⏎ に置換）。"""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


def iter_block_items(parent):
    for child in parent.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)


def cell_text(c):
    t = " ".join(p.text.strip() for p in c.paragraphs if p.text.strip())
    return t.replace("\n", "⏎").replace("\r", "").replace("|", "¦")


src, dst = sys.argv[1], sys.argv[2]
doc = Document(src)
out, ti, pi = [], 0, 0
for blk in iter_block_items(doc):
    if isinstance(blk, Paragraph):
        pi += 1
        t = blk.text.strip().replace("\n", "⏎")
        has = blk._p.findall('.//' + qn('w:drawing'))
        if t:
            out.append(f"[P{pi}]{'[図]' if has else ''}<{blk.style.name}> {t}")
        elif has:
            out.append(f"[P{pi}][図（キャプションなし）]")
    else:
        ti += 1
        out.append(f"--- TABLE {ti} ({len(blk.rows)}行 x {len(blk.columns)}列) ---")
        for r in blk.rows:
            cells, seen = [], set()
            for c in r.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(cell_text(c))
            out.append(" | ".join(cells))
        out.append(f"--- /TABLE {ti} ---")
open(dst, "w", encoding="utf-8").write("\n".join(out))
print(f"{dst}: 段落{pi} 表{ti}")
