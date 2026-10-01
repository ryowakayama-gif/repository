# -*- coding: utf-8 -*-
"""計画素案 Ver.2.6 → Ver.2.7（図を3点加え、採番を振り直す）

町のご指示により、図表データ管理台帳の04シート「不足している図」のうち
**追加の資料を要さず当社の作業のみで作れる3件**を作成し、本文に差し込む。

  図2-4（新）　要介護認定者数と認定率の推移　　　　　→ 2-3
  図3-1（新）　交付金の得点の推移と指標群別の内訳　　→ 3-3
  図6-1（新）　認知症施策のKPIの3層構造　　　　　　→ 6-3

図は本文の出現順に採番するため、既存の2点が繰り下がる。

  図2-4 サービス受給者の区分別構成　→ **図2-5**
  図2-5 サービス区分別の年間給付費　→ **図2-6**

あわせて、本文の参照の誤りを1件直す。

  2-2 の本文が高齢化率の図を「（図2-5）」と参照していた。正しくは**図2-2**。
  図番号を出現順に振り直した際（Ver.2.5）に、本文の参照が追随していなかった。

  python3 07_ソーススクリプト/fix_soan_v270_figures.py

前提：`python3 07_ソーススクリプト/build_plan_figures.py` を先に実行し、
08_図表 に9点の png があること。
"""
import copy
import os
import re
import shutil
import struct
import sys
import zipfile

import docx
from docx.oxml.ns import qn
from docx.shared import Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_zuhyo import ZU                      # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.7_図表追加版.docx"
FIGDIR = "08_図表"
EMU_IN = 914400

# 加える図と、その差し込み位置の見つけ方・幅（インチ）
#   anchor は本文の要素を1つ返す関数。その直後に［図］［キャプション］を置く。
NEW_FIGS = [
    ("図2-4", 4.9),
    ("図3-1", 5.6),
    ("図6-1", 5.6),
]

# 本文の参照の是正（Ver.2.5の採番の振り直しに追随していなかった）
REF_FIX = [("5位でした（図2-5）", "5位でした（図2-2）")]

HEADER_NEW = ("川崎町高齢者保健福祉計画・第10期介護保険事業計画　"
              "計画書素案 Ver.2.7")
VER = {
    4: "計画書素案 Ver.2.7",
    5: "（図表追加版）",
    21: ("⚠ 頁番号は、本素案Ver.2.7の組版によるものです。"
         "図表の差替え・加筆により頁の割付けが動いた場合は、"
         "目次の頁番号も併せて更新します。"),
}
VER_SWAP = [("本素案Ver.2.6", "本素案Ver.2.7"),
            ("計画書素案 Ver.2.6", "計画書素案 Ver.2.7")]


def png_size(b):
    if b[:8] != b"\x89PNG\r\n\x1a\n":
        raise SystemExit("png ではない")
    return struct.unpack(">II", b[16:24])


def cap_of(d):
    return d["title"] + (d["cap_add"] or "")


def find_tbl(doc, pred):
    for t in doc.tables:
        if pred(t):
            return t
    return None


def anchors(doc):
    """新しい図を差し込む位置（この要素の直後に置く）を返す。"""
    body = doc.element.body
    kids = list(body.iterchildren())

    def after_tbl(t, extra=0):
        i = kids.index(t._tbl)
        return kids[i + extra]

    # 図2-4　2-3 の認定率の表 → その直後の出典の段落の後ろ
    t = find_tbl(doc, lambda t: len(t.rows) == 7 and len(t.columns) == 5
                 and "認定率" in t.rows[0].cells[3].text)
    if t is None:
        raise SystemExit("2-3 の認定率の表が見つからない")
    a24 = after_tbl(t, 1)
    if not "".join(n.text or "" for n in a24.iter(qn("w:t"))).startswith("出典"):
        raise SystemExit("認定率の表の次が出典の段落でない")

    # 図3-1　3-3 の得点の推移の表の直後
    t = find_tbl(doc, lambda t: len(t.rows) == 7 and len(t.columns) == 4
                 and "得点" in t.rows[1].cells[0].text)
    if t is None:
        raise SystemExit("3-3 の得点の表が見つからない")
    a31 = t._tbl

    # 図6-1　6-3 のKPIの表の直後
    t = find_tbl(doc, lambda t: len(t.rows) == 15 and len(t.columns) == 5
                 and t.rows[0].cells[0].text.strip() == "層")
    if t is None:
        raise SystemExit("6-3 のKPIの表が見つからない")
    a61 = t._tbl
    return {"図2-4": a24, "図3-1": a31, "図6-1": a61}


def templates(doc):
    """既存の［図の段落］［キャプションの段落］を雛形として取り出す。"""
    body = doc.element.body
    kids = list(body.iterchildren())
    for i, el in enumerate(kids):
        if el.tag == qn("w:p") and el.findall(".//" + qn("w:drawing")):
            cap = kids[i + 1]
            return el, cap
    raise SystemExit("図の段落が見つからない")


def set_el(el, text):
    from fix_soan_v111 import set_el as f
    return f(el, text)


def insert_figure(doc, no, width_in, anchor, p_img_tpl, p_cap_tpl):
    """anchor の直後に［図］［キャプション］の2段落を置く。"""
    d = next(x for x in ZU if x["no"] == no)
    path = os.path.join(FIGDIR, d["png"])
    b = open(path, "rb").read()
    w, h = png_size(b)
    cx = int(round(width_in * EMU_IN))
    cy = int(round(cx * h / w))

    # 図の段落　末尾に作ってから位置を移す
    doc.add_picture(path, width=Emu(cx), height=Emu(cy))
    p_img = doc.paragraphs[-1]._p
    p_img.getparent().remove(p_img)
    # 既存の図の段落の書式（中央揃えなど）を写す
    pPr = p_img_tpl.find(qn("w:pPr"))
    if pPr is not None:
        old = p_img.find(qn("w:pPr"))
        if old is not None:
            p_img.remove(old)
        p_img.insert(0, copy.deepcopy(pPr))
    # add_picture が作る run には書体の指定がないため、雛形の run から写す
    # （書体の点検の検査3「書体の指定がない run がないこと」に当たるため）
    tpl_r = p_img_tpl.find(qn("w:r"))
    rPr = tpl_r.find(qn("w:rPr")) if tpl_r is not None else None
    if rPr is not None:
        for r in p_img.findall(qn("w:r")):
            if r.find(qn("w:rPr")) is None:
                r.insert(0, copy.deepcopy(rPr))

    p_cap = copy.deepcopy(p_cap_tpl)
    set_el(p_cap, "%s　%s" % (d["no"], cap_of(d)))

    anchor.addnext(p_cap)
    anchor.addnext(p_img)
    return cx, cy


def rebuild_index(doc):
    """巻末の図表番号一覧を、実際に差し込んだ図から作り直す。"""
    t = find_tbl(doc, lambda t: [c.text.strip() for c in t.rows[0].cells][:2]
                 == ["図番号", "図表名"])
    if t is None:
        raise SystemExit("図表番号一覧の表が見つからない")
    tpl = copy.deepcopy(t.rows[1]._tr)
    while len(t.rows) - 1 < len(ZU):
        t._tbl.append(copy.deepcopy(tpl))
    while len(t.rows) - 1 > len(ZU):
        t._tbl.remove(t.rows[-1]._tr)
    from fix_soan_v111 import set_cell
    for r, d in zip(t.rows[1:], ZU):
        set_cell(r.cells[0], d["no"])
        set_cell(r.cells[1], d["title"])
        set_cell(r.cells[2], d["sec"])
        set_cell(r.cells[3], d["src"])
    return len(ZU)


def body_embeds(path):
    with zipfile.ZipFile(path) as z:
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8")
    rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="(media/[^"]+)"', rels))
    out, seen = [], set()
    for m in re.finditer(r'r:embed="(rId\d+)"', doc):
        rid = m.group(1)
        if rid in rmap and rid not in seen:
            seen.add(rid)
            out.append((rid, "word/" + rmap[rid]))
    return out


def fix_extents(doc, extents):
    def one(m):
        blk = m.group(0)
        mr = re.search(r'r:embed="(rId\d+)"', blk)
        if not mr or mr.group(1) not in extents:
            return blk
        cx, cy = extents[mr.group(1)]
        blk = re.sub(r'(<wp:extent) cx="\d+" cy="\d+"',
                     r'\1 cx="%d" cy="%d"' % (cx, cy), blk)
        blk = re.sub(r'(<a:ext) cx="\d+" cy="\d+"',
                     r'\1 cx="%d" cy="%d"' % (cx, cy), blk)
        return blk
    return re.sub(r"<w:drawing>.*?</w:drawing>", one, doc, flags=re.S)


def swap_media(path, mapping, extents):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in mapping:
                data = mapping[item.filename]
            elif item.filename == "word/document.xml":
                data = fix_extents(data.decode("utf-8"),
                                   extents).encode("utf-8")
            elif item.filename == "word/header1.xml":
                s = re.sub(r"計画書素案 ?Ver\.[0-9.]+",
                           "計画書素案 Ver.2.7", data.decode("utf-8"))
                data = s.encode("utf-8")
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    shutil.move(tmp, path)


def main():
    if not os.path.exists(SRC):
        raise SystemExit("素案が見つからない：" + SRC)
    for d in ZU:
        p = os.path.join(FIGDIR, d["png"])
        if not os.path.exists(p):
            raise SystemExit(
                "画像がない：%s（先に build_plan_figures.py を実行する）" % p)

    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body

    # ── ① 図を3点加える
    p_img_tpl, p_cap_tpl = templates(doc)
    anc = anchors(doc)
    added = []
    for no, w in NEW_FIGS:
        cx, cy = insert_figure(doc, no, w, anc[no], p_img_tpl, p_cap_tpl)
        added.append((no, w, cx, cy))

    # ── ② キャプションを data_zuhyo.py に合わせる（繰り下げた2点を含む）
    want = {d["no"]: "%s　%s" % (d["no"], cap_of(d)) for d in ZU}
    caps = []
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        if re.match(r"^図\d+-\d+　", t):
            caps.append(el)
    if len(caps) != len(ZU):
        raise SystemExit("キャプションが%d件（図は%d点）" % (len(caps), len(ZU)))
    n_cap = 0
    for el, d in zip(caps, ZU):
        t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        if t != want[d["no"]]:
            set_el(el, want[d["no"]])
            n_cap += 1

    # ── ③ 本文の参照の是正と版表記
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    for i, txt in VER.items():
        set_el(ps[i - 1], txt)
    n_ref = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        new = t
        for a, b in REF_FIX + VER_SWAP:
            new = new.replace(a, b)
        if new != t:
            set_el(el, new)
            n_ref += 1

    # ── ④ 巻末の図表番号一覧
    if n_ref < 1 + len(REF_FIX) - 1:
        raise SystemExit("本文の参照の是正が当たっていない（%d件）" % n_ref)
    n_idx = rebuild_index(doc)
    doc.save(DST)

    # ── ⑤ 画像の中身と表示の大きさを data_zuhyo.py に合わせる
    embeds = body_embeds(DST)
    if len(embeds) != len(ZU):
        raise SystemExit("素案の画像 %d 点と図の定義 %d 点が合わない"
                         % (len(embeds), len(ZU)))
    with zipfile.ZipFile(DST) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")
    old = {}
    for m in re.finditer(r"<w:drawing>.*?</w:drawing>", doc_xml, flags=re.S):
        mr = re.search(r'r:embed="(rId\d+)"', m.group(0))
        me = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', m.group(0))
        if mr and me:
            old[mr.group(1)] = (int(me.group(1)), int(me.group(2)))
    mapping, extents, log = {}, {}, []
    for d, (rid, part) in zip(ZU, embeds):
        b = open(os.path.join(FIGDIR, d["png"]), "rb").read()
        w, h = png_size(b)
        cx, cy_old = old[rid]
        cy = int(round(cx * h / w))
        mapping[part] = b
        extents[rid] = (cx, cy)
        log.append((d["no"], d["png"], len(b), cx, cy))
    swap_media(DST, mapping, extents)

    print("保存：", DST)
    for no, w, cx, cy in added:
        print("  ＋%s を差し込みました（幅 %.1f インチ）" % (no, w))
    for no, png, size, cx, cy in log:
        print("  %s ← %-22s %7d バイト　幅 %d／高さ %d" % (no, png, size, cx, cy))
    print("  キャプションの是正 %d 件／本文の参照・版表記の是正 %d 件"
          % (n_cap, n_ref))
    print("  巻末の図表番号一覧 %d 行" % n_idx)


if __name__ == "__main__":
    main()
