# -*- coding: utf-8 -*-
"""提供資料を本文（docx）へ反映する。

 ・下水道：p.3 処理区域図の貼込、p.19 管渠の本文とグラフ、p.20 処理場の概要表、p.25 連動修正
 ・簡易水道：p.3 給水区域図の貼込

OOXMLを直接編集し、EMF図表・コメント・書式を壊さない。
"""
import zipfile, shutil, sys, re, os
from lxml import etree
from PIL import Image

W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A  = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
EMU = 914400
log = []


def replace_in_para(p, old, new):
    """run分割に対応した置換。最初のrunの書式を引き継ぐ。"""
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


class Doc:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.parts = {n: self.z.read(n) for n in self.z.namelist()}
        self.infos = self.z.infolist()
        self.root = etree.fromstring(self.parts['word/document.xml'])
        self.ps = list(self.root.iter(W + 'p'))
        rx = self.parts['word/_rels/document.xml.rels'].decode()
        self.next_rid = max(int(m) for m in re.findall(r'Id="rId(\d+)"', rx)) + 1
        self.next_img = max([int(m) for m in re.findall(r'image(\d+)\.', ' '.join(self.parts))] or [0]) + 1
        self.docpr = 9000001

    def add_image(self, png):
        """PNGをmediaに追加し、rIdを返す。"""
        name = f'image{self.next_img}.png'; self.next_img += 1
        self.parts[f'word/media/{name}'] = open(png, 'rb').read()
        rid = f'rId{self.next_rid}'; self.next_rid += 1
        rels = etree.fromstring(self.parts['word/_rels/document.xml.rels'])
        etree.SubElement(rels, f'{{{RELNS}}}Relationship',
                         Id=rid, Type=IMG_REL, Target=f'media/{name}')
        self.parts['word/_rels/document.xml.rels'] = etree.tostring(
            rels, xml_declaration=True, encoding='UTF-8', standalone=True)
        return rid, name

    def _drawing_xml(self, rid, name, cx, cy):
        self.docpr += 1
        return f'''<w:r xmlns:w="{W[1:-1]}"><w:rPr><w:noProof/></w:rPr><w:drawing xmlns:wp="{WP[1:-1]}">
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="{self.docpr}" name="{name}"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{A[1:-1]}" noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic xmlns:a="{A[1:-1]}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:nvPicPr><pic:cNvPr id="{self.docpr}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip xmlns:r="{R[1:-1]}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'''

    def insert_image(self, pidx, png, width_in, why, center=True):
        im = Image.open(png)
        cx = int(width_in * EMU)
        cy = int(width_in * im.height / im.width * EMU)
        rid, name = self.add_image(png)
        p = self.ps[pidx]
        if center:
            self._set_jc(p, 'center')
        p.append(etree.fromstring(self._drawing_xml(rid, name, cx, cy)))
        log.append(('挿入', f'段落{pidx}', f'{name} {cx/EMU:.2f}×{cy/EMU:.2f}in', why))

    # w:pPr の子要素はスキーマで順序が決まっているため、正しい位置へ入れる
    PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
                 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
                 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
                 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
                 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
                 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
                 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
                 'pPrChange']

    def _set_jc(self, p, val):
        pPr = p.find(W + 'pPr')
        if pPr is None:
            pPr = etree.Element(W + 'pPr'); p.insert(0, pPr)
        jc = pPr.find(W + 'jc')
        if jc is not None:
            jc.set(W + 'val', val); return
        jc = etree.Element(W + 'jc'); jc.set(W + 'val', val)
        want = self.PPR_ORDER.index('jc')
        pos = len(pPr)
        for i, ch in enumerate(pPr):
            nm = etree.QName(ch).localname
            if nm in self.PPR_ORDER and self.PPR_ORDER.index(nm) > want:
                pos = i; break
        pPr.insert(pos, jc)

    def swap_image(self, pidx, png, width_in=None, why=''):
        """既存の図の参照先を差し替える。width_in省略時は既存の幅を維持。"""
        p = self.ps[pidx]
        dr = next(p.iter(W + 'drawing'))
        ext = dr.find('.//' + WP + 'extent')
        blip = dr.find('.//' + A + 'blip')
        old_cx, old_cy = int(ext.get('cx')), int(ext.get('cy'))
        im = Image.open(png)
        cx = int((width_in or old_cx / EMU) * EMU)
        cy = int(cx * im.height / im.width)
        rid, name = self.add_image(png)
        blip.set(R + 'embed', rid)
        ext.set('cx', str(cx)); ext.set('cy', str(cy))
        # a:xfrm 配下の a:ext のみ（a:extLst の拡張要素と取り違えない）
        for xf in dr.iter(A + 'xfrm'):
            for a_ext in xf.findall(A + 'ext'):
                a_ext.set('cx', str(cx)); a_ext.set('cy', str(cy))
        log.append(('差替', f'段落{pidx}', f'{old_cx/EMU:.2f}×{old_cy/EMU:.2f}in → {name} {cx/EMU:.2f}×{cy/EMU:.2f}in', why))

    def edit(self, pidx, old, new, why):
        if not replace_in_para(self.ps[pidx], old, new):
            raise SystemExit(f'!! 段落{pidx} に該当文字列なし: {old[:30]}')
        log.append(('本文', f'段落{pidx}', f'{old[:26]}… → {new[:26]}…', why))

    def save(self, dst):
        self.parts['word/document.xml'] = etree.tostring(
            self.root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
        done = set()
        for it in self.infos:
            zo.writestr(it, self.parts[it.filename]); done.add(it.filename)
        for n, b in self.parts.items():
            if n not in done:
                zo.writestr(n, b)
        zo.close()


# ══════════════ 下水道 ══════════════
G = Doc(sys.argv[1])
G.edit(837,
       '令和５（20２３）年度末の時点で、下水道事業の総管渠延長は、１７１,３４７.05ｍ（管更生含む）に達し、村内での下水道普及率はほぼ１００％となっています。',
       '令和５（２０２３）年度末の時点で、下水道事業の総管渠延長は、２５，２７０．６６ｍとなっています。村内での下水道普及率は６９．５％（令和５年度決算）であり、処理区域内の水洗化率は９８．５％と高い水準にあります。',
       'p.19 提供資料「管渠延長.pdf」の合計値へ。普及率は経営比較分析表の値')
G.edit(838,
       '本計画期間中の１０年間で耐用年数を迎える管渠は全体の約２６％の昭和世代に布設された笹尾・城山地区のものです。今後の管渠の更新に向けて長期的な期間で管渠の長寿命化や更新に向けた対策が必要となっています。',
       '管渠は平成４（１９９２）年度以降に布設されており、最も古いもので令和７（２０２５）年度時点において３３年が経過しています。法定耐用年数（５０年）に達するのは令和２４（２０４２）年度以降となるため、本計画期間中に耐用年数を迎える管渠はありませんが、布設から３０年を超える管渠が全体の約４１％（１０，４７７ｍ）を占めることから、長期的な期間で管渠の長寿命化や更新に向けた対策が必要となっています。',
       'p.19 他団体の地名・年代を削除し、提供資料の布設年度に基づく記述へ')
G.edit(900, '管渠は古いもので約４０年が過ぎています。', '管渠は古いもので３０年を超えています。',
       'p.25 最古は平成4年度布設（令和7年度時点で33年）')
G.edit(901,
       '計画期間中に全体の約１／４が耐用年数を迎えることから、更新に向けた資金の確保や長寿命化対策による対応が求められています。',
       '計画期間中に耐用年数を迎える管渠はありませんが、布設から３０年を超える管渠が全体の約４割を占めることから、更新に向けた資金の確保や長寿命化対策による計画的な対応が求められています。',
       'p.25 耐用年数到達の再計算結果を反映')
G.insert_image(416, 'img/gesui_kuiki.png', 6.40, 'p.3 提供資料「区域.pdf」を貼込')
G.swap_image(840, 'img/kankyo_chart.png', 6.53, 'p.19 提供資料の年度別延長で再作成')
G.swap_image(867, 'img/shorijo_table.png', 6.53, 'p.20 提供資料「施設の概要.pdf」の諸元へ差替')
G.save(sys.argv[2])

# ══════════════ 簡易水道 ══════════════
K = Doc(sys.argv[3])
K.insert_image(745, 'img/kansui_kuiki.png', 6.40, 'p.3 提供資料「簡易水道給水区域図」を貼込')
K.save(sys.argv[4])

print(f'=== 適用 {len(log)} 件 ===')
for kind, where, what, why in log:
    print(f'  [{kind}] {where:8s} {what}')
    print(f'            {why}')
