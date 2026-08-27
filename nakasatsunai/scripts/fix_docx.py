# -*- coding: utf-8 -*-
"""中札内村下水道事業経営戦略（案）の確定分修正をdocxへ適用する。
OOXMLを直接編集し、EMF図表・コメント・書式を一切壊さない。"""
import zipfile, shutil, sys
from lxml import etree

W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
SRC = 'work/senryaku.docx'
DST = 'work/senryaku_fixed.docx'

def replace_in_para(p, old, new):
    """段落をまたぐrun分割に対応した置換。最初のrunの書式を引き継ぐ。"""
    ts = list(p.iter(W + 't'))
    texts = [t.text or '' for t in ts]
    full = ''.join(texts)
    idx = full.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    spans, pos = [], 0
    for t, tx in zip(ts, texts):
        spans.append((pos, pos + len(tx), t, tx))
        pos += len(tx)
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

# (段落index, 変更前, 変更後, 根拠)
EDITS = [
    (16,  '建設課', '施設課', '表紙 発行課名／朱書き'),
    (390, '町は札内川流域', '村は札内川流域', 'p.2 朱書き漏れ'),
    (391, '平成９（１９９７）年4月1日に供用を開始',
          '平成９（１９９７）年３月に供用を開始', 'p.2 公表版「平成8年度」と整合'),
    (391, '処理場は1箇所で、町の中心部', '処理場は1箇所で、村の中心部', 'p.2 朱書き'),
    (392, 'これまでの整備により、町全体', 'これまでの整備により、村全体', 'p.2 朱書き'),
    (392, '令和５（２０２３）年４月からは地方公営企業法の全部適用へ移行し',
          '令和４（２０２２）年４月からは地方公営企業法の一部適用（財務規定等の適用）へ移行し',
          'p.2 朱書き＋論点A（公表版は一部適用）'),
    (864, '本町における処理場', '本村における処理場', 'p.20 朱書き漏れ'),
    (864, '平成２２（２０１０）年度から供用開始し、令和６（２０２４）年度現在で１５年の稼働',
          '平成８年度（平成９年３月）から供用開始し、令和６（２０２４）年度現在で２８年の稼働',
          'p.20 公表版「供用開始年度＝平成8年度（28年）」'),
    (1028, '国（県）補助金', '国（道）補助金', 'p.26 朱書き'),
    (1220, '●国（県）補助金', '●国（道）補助金', 'p.34 朱書き'),
    (1221, '国庫（県）補助対象事業', '国庫（道）補助対象事業', 'p.34 朱書き'),
    (1338, '町ホームページ', '村ホームページ', 'p.38 朱書き'),
]

zin = zipfile.ZipFile(SRC)
doc = etree.fromstring(zin.read('word/document.xml'))
paras = list(doc.iter(W + 'p'))

applied, failed = [], []
for i, old, new, why in EDITS:
    if replace_in_para(paras[i], old, new):
        applied.append((i, old, new, why))
    else:
        failed.append((i, old, why))

new_doc = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)

# SmartArt組織図（p.10）の「町長」
diagrams = {}
for name in ('word/diagrams/data1.xml', 'word/diagrams/drawing1.xml'):
    x = zin.read(name).decode('utf-8')
    if '町長' in x:
        diagrams[name] = x.replace('>町長<', '>村長<').encode('utf-8')
        applied.append(('図', '町長', '村長', 'p.10 組織図（SmartArt）／朱書き'))

repl = {'word/document.xml': new_doc, **diagrams}
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    zout.writestr(item, repl.get(item.filename, zin.read(item.filename)))
zout.close(); zin.close()

print('=== 適用した修正 ===')
for i, old, new, why in applied:
    print(f'  [{i}] {old}  →  {new}      ({why})')
print(f'\n適用 {len(applied)} 件 / 失敗 {len(failed)} 件')
for f in failed:
    print('  !! 未適用:', f)
sys.exit(1 if failed else 0)
