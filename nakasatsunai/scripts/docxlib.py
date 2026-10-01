# -*- coding: utf-8 -*-
"""Word（docx）をOOXMLのまま編集するための共通処理。

apply_materials.py で使った Doc クラスを、後続のスクリプトからも使えるよう切り出したもの。
EMF図表・コメント・書式を壊さずに、本文の文字列置換と図の挿入・差替えを行う。
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

    def replace_all(self, pidx, old, new, why):
        """段落内のすべての出現を置換する（テキストボックスと本文の重複に対応）。"""
        n = 0
        while replace_in_para(self.ps[pidx], old, new):
            n += 1
            if n > 10:
                break
        if n:
            log.append(('本文', f'段落{pidx}', f'{old[:22]}… → {new[:22]}…（{n}箇所）', why))
        return n

    def fit_textbox(self, pidx, width_in, why=''):
        """浮動テキストボックスの左位置と幅を本文幅に合わせる。

        元のレイアウトは横置きを前提に幅9.88in・左オフセット−4.64inで配置されており、
        縦置き（本文幅6.54in）のページでは用紙の外へはみ出す。
        """
        cx = int(width_in * EMU)
        hit = 0
        for dr in self.ps[pidx].iter(W + 'drawing'):
            anc = dr.find(WP + 'anchor')
            if anc is None or dr.find('.//' + W + 'txbxContent') is None:
                continue
            ph = anc.find(WP + 'positionH')
            if ph is not None:
                ph.set('relativeFrom', 'margin')
                for ch in list(ph):
                    ph.remove(ch)
                off = etree.SubElement(ph, WP + 'posOffset')
                off.text = '0'
            ext = anc.find(WP + 'extent')
            old_cx = int(ext.get('cx'))
            ext.set('cx', str(cx))
            for xf in anc.iter(A + 'xfrm'):
                for a_ext in xf.findall(A + 'ext'):
                    a_ext.set('cx', str(cx))
            hit += 1
            log.append(('図形', f'段落{pidx}',
                        f'テキストボックス 幅{old_cx/EMU:.2f}in・左−4.64in → 幅{cx/EMU:.2f}in・左0',
                        why))
        return hit

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
