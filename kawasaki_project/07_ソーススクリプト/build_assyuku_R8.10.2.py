# -*- coding: utf-8 -*-
"""本文の圧縮案（令和8年10月2日）

08_作業順位 の順位5。確認事項No.83（計画書の頁数）のご判断の材料として、
**実際に移した版を作って頁数を数える。**

仕様書8の成果品「計画書」は100頁程度である。素案 Ver.2.8 は本文104頁・
資料編13頁で計117頁であり、17％上回っている。

移す候補（詳細な分析であり、計画の本文として読み通す必要が薄いもの）

  3-3　交付金の枝番の分析　　5つの小見出し
  9-3　感度分析・乗率・収納率　3つの小見出し
  6-4　追加候補（協議用）　　2つの小見出し

**6-4 の認知症施策の体系（6節8項28事業）は動かさない。**
認知症施策推進計画の必須記載事項であり、本文に置く必要がある。

  出力　01_第10期_最新版成果品/川崎町_計画書素案_v2.8c_圧縮案_R8.10.2.docx
        01_第10期_最新版成果品/川崎町_計画書_別冊_詳細分析_R8.10.2.docx

**これは案であり、ご判断をいただくまで Ver.2.8 を正本とする。**
"""
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile

import docx
import pypdf
from docx.oxml.ns import qn

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.8_第9期対比版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.8c_圧縮案_R8.10.2.docx"
BESSATSU = "01_第10期_最新版成果品/川崎町_計画書_別冊_詳細分析_R8.10.2.docx"

# （節, 小見出しの冒頭, 別冊での章立て）
MOVE = [
    ("3-3", "▌ 規模区分・過疎地域該当ごとの平均（参考）", "1 交付金の評価の詳細"),
    ("3-3", "▌ 3年間で変わらない強みと、一度も取れていない指標",
     "1 交付金の評価の詳細"),
    ("3-3", "▌ 失点を取り戻す順序", "1 交付金の評価の詳細"),
    ("3-3", "▌ 体制・取組指標群の失点は3か所・27点", "1 交付金の評価の詳細"),
    ("3-3", "▌ 成果指向型配分枠（配点100点）", "1 交付金の評価の詳細"),
    ("9-3", "▌ 前提を動かした場合の幅", "2 保険料算定の感度分析"),
    ("9-3", "▌ 所得段階の乗率の検討（現行13段階）", "2 保険料算定の感度分析"),
    ("9-3", "▌ 予定収納率の設定", "2 保険料算定の感度分析"),
    ("6-4", "▌ 第2回策定委員会でご判断いただく事業の追加候補（15件）",
     "3 認知症施策の追加候補（協議用）"),
    ("6-4", "▌ 追加候補の見込量・費用・財源の枠組み",
     "3 認知症施策の追加候補（協議用）"),
]

LEAD = ("本別冊は、川崎町高齢者保健福祉計画・第10期介護保険事業計画の"
        "本文に掲げた分析のうち、詳細にわたる部分をまとめたものです。"
        "本文の 3-3（評価指標の達成状況）・9-3（介護保険料の算定）・"
        "6-4（認知症施策の体系）に対応します。"
        "計画の内容は本文に掲げており、本別冊はその根拠を示すものです。")


def txt(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()


def set_el(el, text):
    rs = el.findall(qn("w:r"))
    for r in rs[1:]:
        el.remove(r)
    ts = rs[0].findall(qn("w:t"))
    for t in ts[1:]:
        rs[0].remove(t)
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def blocks(doc):
    """節ごと・小見出しごとに、本文の要素の並びを切り出す。"""
    kids = [el for el in doc.element.body.iterchildren()
            if el.tag in (qn("w:p"), qn("w:tbl"))]
    cur = sub = None
    out = {}
    for el in kids:
        if el.tag == qn("w:p"):
            s = txt(el)
            m = re.match(r"^(\d+-\d+)[　 ]", s)
            if m:
                cur, sub = m.group(1), None
            elif s.startswith("▌"):
                sub = s
        if cur and sub:
            out.setdefault((cur, sub), []).append(el)
    return out


def pages(path):
    with tempfile.TemporaryDirectory() as wd:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                        "--outdir", wd, os.path.abspath(path)],
                       check=True, env=dict(os.environ, HOME=wd),
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        f = [x for x in os.listdir(wd) if x.endswith(".pdf")][0]
        return len(pypdf.PdfReader(os.path.join(wd, f)).pages)


def main():
    before = pages(SRC)
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    bl = blocks(doc)

    # 小見出しの文言は実物から照合する（前方一致）
    picked, miss = [], []
    for sec, head, cap in MOVE:
        key = next((k for k in bl
                    if k[0] == sec and k[1].startswith(head)), None)
        if key is None:
            miss.append(f"{sec}／{head}")
            continue
        picked.append((sec, key, cap))
    if miss:
        for m in miss:
            print("  × 見つからない小見出し：", m)
        sys.exit(1)

    # ══════════════════════════ 別冊を作る（素案の書式を引き継ぐ）
    shutil.copy(SRC, BESSATSU)
    bdoc = docx.Document(BESSATSU)
    bbody = bdoc.element.body
    tpl_sho = tpl_honbun = None
    for el in bbody.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = txt(el)
        if tpl_sho is None and s.startswith("▌"):
            tpl_sho = copy.deepcopy(el)
        if tpl_honbun is None and s.startswith("令和7年度から令和8年度への"):
            tpl_honbun = copy.deepcopy(el)
    sectPr = bbody.find(qn("w:sectPr"))
    for el in list(bbody.iterchildren()):
        if el is not sectPr:
            bbody.remove(el)

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            bbody.append(el)

    p = copy.deepcopy(tpl_sho)
    set_el(p, "川崎町高齢者保健福祉計画・第10期介護保険事業計画　"
              "別冊　詳細分析")
    add(p)
    p = copy.deepcopy(tpl_honbun)
    set_el(p, LEAD)
    add(p)

    n_moved = 0
    last_cap = None
    for sec, key, cap in picked:
        if cap != last_cap:
            p = copy.deepcopy(tpl_sho)
            set_el(p, cap)
            add(p)
            last_cap = cap
        for el in bl[key]:
            add(copy.deepcopy(el))
            n_moved += 1
    bdoc.save(BESSATSU)

    # ══════════════════════════ 本文から取り除く
    body = doc.element.body
    removed = 0
    for sec, key, cap in picked:
        for el in bl[key]:
            if el.getparent() is not None:
                body.remove(el)
                removed += 1
    # 移した先への案内を 3-3・9-3・6-4 の末尾に置く
    GUIDE = {
        "3-3": "⚠ 交付金の評価の詳細（規模区分別の平均、3年連続で満点の指標、"
               "失点を取り戻す順序、体制・取組指標群の失点3か所の明細、"
               "成果指向型配分枠）は、別冊「詳細分析」1に掲げています。",
        "9-3": "⚠ 前提を動かした場合の幅（感度分析）、所得段階の乗率の検討、"
               "予定収納率の設定は、別冊「詳細分析」2に掲げています。",
        "6-4": "⚠ 第2回策定委員会でご判断いただく事業の追加候補（15件）と、"
               "その見込量・費用・財源の枠組みは、"
               "別冊「詳細分析」3に掲げています。",
    }
    kids = list(body.iterchildren())
    tpl = None
    for el in kids:
        if el.tag == qn("w:p") and txt(el).startswith("令和7年度から令和8年度への"):
            tpl = el
            break
    for sec, g in GUIDE.items():
        # その節の最後の要素を探す
        cur = None
        last = None
        for el in body.iterchildren():
            if el.tag == qn("w:p"):
                m = re.match(r"^(\d+-\d+)[　 ]", txt(el))
                if m:
                    cur = m.group(1)
            if cur == sec:
                last = el
        if last is None:
            continue
        p = copy.deepcopy(tpl)
        set_el(p, g)
        last.addnext(p)
    doc.save(DST)

    after = pages(DST)
    bp = pages(BESSATSU)
    print("本文の圧縮案")
    print(f"  移した小見出し {len(picked)}／要素 {removed}")
    print(f"  本文　Ver.2.8 {before}頁 → 圧縮案 {after}頁"
          f"（{after - before:+d}頁）")
    print(f"  別冊「詳細分析」{bp}頁")
    print(f"  計画書（本文＋資料編13頁）"
          f"　現行 {before + 13}頁 → 圧縮案 {after + 13}頁")
    print("  仕様書8の「100頁程度」との差"
          f"　現行 {before + 13 - 100:+d}頁 → 圧縮案 {after + 13 - 100:+d}頁")
    if after >= before:
        print("  × 圧縮になっていない")
        sys.exit(1)


if __name__ == "__main__":
    main()
