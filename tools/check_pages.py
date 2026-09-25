import sys, pymupdf, re
from docx import Document
pdf, docx = sys.argv[1], sys.argv[2]
d = pymupdf.open(pdf)
norm = lambda s: re.sub(r"\s|　", "", s)
pages = [norm(p.get_text()) for p in d]
def pageof(s, frm=1):
    s = norm(s)
    if not s: return None
    for i in range(frm - 1, len(pages)):
        if s in pages[i]: return i + 1
    return None
doc = Document(docx)
capmap = {}
for p in doc.paragraphs:
    m = re.match(r"表(\d+)　", p.text)
    if m: capmap[int(m.group(1))] = p.text
split = []
for n, t in enumerate(doc.tables, 1):
    cap = capmap.get(n)
    if not cap: continue
    pc = pageof(cap)
    last = "".join(c.text for c in t.rows[-1].cells)[:60]
    pl = pageof(last, pc or 1)
    if pc and pl and pc != pl:
        split.append((n, pc, pl, cap[:28]))
print("総ページ %d" % len(pages))
# 表紙（1ページ目）と最終ページは短くて構わない
_short = [(i, len(t)) for i, t in enumerate(pages, 1)
          if len(t) < 400 and 1 < i < len(pages)]
print("中身の少ないページ（表紙・最終ページを除く）:", _short)
print("紙面をまたぐ表 %d件" % len(split))
for x in split: print("   表%d 見出しp%d→最終行p%d  %s" % x)
