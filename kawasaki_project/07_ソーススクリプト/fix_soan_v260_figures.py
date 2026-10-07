# -*- coding: utf-8 -*-
"""計画素案 Ver.2.5 → Ver.2.6（図の差替えと図表の整理）

Ver.2.5 に差し込まれていた6点の画像は、図の数値が
`create_charts.py`・`make_charts_v114/v115/v116.py` の4つに散らばっていた
ころの出力であり、次の食い違いを抱えていた。

  ① 図9-1 の3パターンが 6,822／6,170／5,518 のままであった
     （素案 Ver.2.3 で 6,822／6,143／5,464 に是正したが図は追随していない）
  ② 画像に焼き込んだ図番号が素案 Ver.2.5 の採番と4点食い違っていた
     画像内「図2-5」＝素案の図2-2、画像内「図2-4」＝素案の図2-3、
     画像内「図2-2」＝素案の図2-4、画像内「図2-3」＝素案の図2-5

本スクリプトは、図の数値の正本 `data_zuhyo.py` から作り直した
`08_図表/fig*.png`（`build_plan_figures.py` の出力）で6点を差し替え、
キャプションを `title＋cap_add` に揃え、版表記を Ver.2.6 に改める。

画像は docx の中の同じパート名のまま中身だけを入れ替える。
本文の差込み位置・回り込み・段落は一切動かさない。
**幅（cx）は Ver.2.5 のまま据え置き、高さ（cy）だけを
新しい画像の縦横比に合わせて求め直す**（縦横比の崩れを避けるため）。

  python3 07_ソーススクリプト/fix_soan_v260_figures.py

前提：`python3 07_ソーススクリプト/build_plan_figures.py` を先に実行し、
08_図表 に6点の png があること。
"""
import os
import re
import shutil
import struct
import sys
import zipfile

import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_zuhyo import ZU                      # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.5_書体統一版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx"
FIGDIR = "08_図表"

HEADER_NEW = ("川崎町高齢者保健福祉計画・第10期介護保険事業計画　"
              "計画書素案 Ver.2.6")

# ══════════════════════════ 版表記（段落番号は Ver.2.5 と同じ）
VER = {
    4: "計画書素案 Ver.2.6",
    5: "（図表整理版）",
    21: ("⚠ 頁番号は、本素案Ver.2.6の組版によるものです。"
         "図表の差替え・加筆により頁の割付けが動いた場合は、"
         "目次の頁番号も併せて更新します。"),
}
VER_SWAP = [("本素案Ver.2.5", "本素案Ver.2.6"),
            ("計画書素案 Ver.2.5", "計画書素案 Ver.2.6")]


def png_size(b):
    """png のバイト列から画素の幅・高さを読む（IHDR）。"""
    if b[:8] != b"\x89PNG\r\n\x1a\n":
        raise SystemExit("png ではない")
    w, h = struct.unpack(">II", b[16:24])
    return w, h


def body_embeds(path):
    """本文に現れる順に (rId, パート名) を返す。"""
    with zipfile.ZipFile(path) as z:
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8")
    rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="(media/[^"]+)"', rels))
    out = []
    for m in re.finditer(r'r:embed="(rId\d+)"', doc):
        rid = m.group(1)
        if rid in rmap and rid not in [x[0] for x in out]:
            out.append((rid, "word/" + rmap[rid]))
    return out


def swap_media(path, mapping, extents):
    """画像の中身を入れ替え、あわせて表示の大きさ（cx・cy）を書き直す。

    mapping  パート名 → 差し込む png の中身（バイト列）
    extents  rId → (cx, cy)
    """
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
                s = data.decode("utf-8")
                s = re.sub(r"計画書素案 ?Ver\.[0-9.]+",
                           "計画書素案 Ver.2.6", s)
                data = s.encode("utf-8")
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    shutil.move(tmp, path)


def fix_extents(doc, extents):
    """drawing ごとの wp:extent と a:ext を新しい大きさに書き直す。

    1つの drawing の中に wp:extent（枠の大きさ）と a:ext（図形の大きさ）が
    あり、双方を揃えないと画像だけが枠からはみ出す。
    r:embed は drawing の末尾側にあるため、drawing 単位で切って扱う。
    """
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


def cap_of(d):
    return d["title"] + (d["cap_add"] or "")


def main():
    if not os.path.exists(SRC):
        raise SystemExit("素案が見つからない：" + SRC)
    for d in ZU:
        p = os.path.join(FIGDIR, d["png"])
        if not os.path.exists(p):
            raise SystemExit(
                "画像がない：%s（先に build_plan_figures.py を実行する）" % p)

    embeds = body_embeds(SRC)
    if len(embeds) != len(ZU):
        raise SystemExit("素案の画像 %d 点と図の定義 %d 点が合わない"
                         % (len(embeds), len(ZU)))

    # 本文に現れる順＝ZU の並び（図2-1→2-2→2-3→2-4→2-5→9-1）
    with zipfile.ZipFile(SRC) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")
    old_cx = {}
    for m in re.finditer(r"<w:drawing>.*?</w:drawing>", doc_xml, flags=re.S):
        blk = m.group(0)
        mr = re.search(r'r:embed="(rId\d+)"', blk)
        me = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', blk)
        if mr and me:
            old_cx[mr.group(1)] = (int(me.group(1)), int(me.group(2)))

    shutil.copy(SRC, DST)

    mapping, extents, log = {}, {}, []
    for d, (rid, part) in zip(ZU, embeds):
        b = open(os.path.join(FIGDIR, d["png"]), "rb").read()
        w, h = png_size(b)
        cx, cy_old = old_cx[rid]
        cy = int(round(cx * h / w))
        mapping[part] = b
        extents[rid] = (cx, cy)
        log.append((d["no"], d["png"], len(b), cx, cy_old, cy))

    swap_media(DST, mapping, extents)

    # ── キャプションと版表記
    doc = docx.Document(DST)
    body = doc.element.body
    from fix_soan_v111 import set_el
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    for i, txt in VER.items():
        set_el(ps[i - 1], txt)

    want = {d["no"]: cap_of(d) for d in ZU}
    n_cap = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        m = re.match(r"^(図\d+-\d+)　(.+)$", t.strip())
        if m and m.group(1) in want:
            new = "%s　%s" % (m.group(1), want[m.group(1)])
            if new != t.strip():
                set_el(el, new)
                n_cap += 1
            continue
        new = t
        for a, b2 in VER_SWAP:
            new = new.replace(a, b2)
        if new != t:
            set_el(el, new)
    doc.save(DST)

    print("保存：", DST)
    for no, png, size, cx, cy_old, cy in log:
        print("  %s ← %-20s %7d バイト　幅 %d 据置き／高さ %d → %d"
              % (no, png, size, cx, cy_old, cy))
    print("  キャプションの是正 %d 件" % n_cap)


if __name__ == "__main__":
    main()
