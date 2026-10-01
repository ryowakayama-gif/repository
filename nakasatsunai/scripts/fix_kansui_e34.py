# -*- coding: utf-8 -*-
"""簡易水道本文へ E-3・E-4 を反映する（下水道側と同様の是正）。
OOXMLを直接編集し、図表・書式を壊さない。"""
import zipfile, sys
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
SRC, DST = sys.argv[1], sys.argv[2]


def replace_in_para(p, old, new):
    ts = list(p.iter(W + 't'))
    texts = [t.text or '' for t in ts]
    full = ''.join(texts)
    idx = full.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    spans, pos = [], 0
    for t, tx in zip(ts, texts):
        spans.append((pos, pos + len(tx), t, tx)); pos += len(tx)
    first = True
    for s, e, t, tx in spans:
        if e <= idx or s >= end:
            continue
        a, b = max(idx, s) - s, min(end, e) - s
        t.text = tx[:a] + (new if first else '') + tx[b:]
        first = False
        if t.text != t.text.strip():
            t.set(XS, 'preserve')
    return True


EDITS = [
    (972, '（下水熱・下水汚泥・発電等）', '（小水力発電等）',
     'E-4　資産活用の状況・エネルギー利用の例示を簡易水道に即した内容へ'),
    (1400, '●国（県）補助金', '●国（道）補助金',
     'E-3　下水道側の朱書き指示と同様に「県」→「道」'),
    (1401, '国庫（県）補助対象事業', '国庫（道）補助対象事業',
     'E-3　同上'),
]

z = zipfile.ZipFile(SRC)
root = etree.fromstring(z.read('word/document.xml'))
ps = list(root.iter(W + 'p'))

applied, failed = [], []
for i, old, new, why in EDITS:
    (applied if replace_in_para(ps[i], old, new) else failed).append((i, old, new, why))

parts = {n: z.read(n) for n in z.namelist()}
parts['word/document.xml'] = etree.tostring(root, xml_declaration=True,
                                            encoding='UTF-8', standalone=True)
zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    zo.writestr(it, parts[it.filename])
zo.close(); z.close()

print(f'=== 適用 {len(applied)} 件 / 失敗 {len(failed)} 件 ===')
for i, old, new, why in applied:
    print(f'  [段落{i}] {old} → {new}')
    print(f'            {why}')
for f in failed:
    print('  !! 未適用:', f)
sys.exit(1 if failed else 0)
