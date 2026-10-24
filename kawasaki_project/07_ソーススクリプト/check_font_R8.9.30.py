# -*- coding: utf-8 -*-
"""書体と文字の大きさの点検

大雪地区広域連合の協議用素案（令和8年8月）の書式に合わせているかを確かめる。

  検査1　既定の段落スタイル（Normal）が定義されているか
  　　　 Title・Heading1〜6 はいずれも basedOn="Normal" を持つため、
  　　　 Normal がないと Word が自前の「標準」（明朝体のことがある）を当てる
  検査2　すべてのスタイルに HG丸ｺﾞｼｯｸM-PRO が明示されているか（R8.10.23 游ゴシックから改め）
  検査3　書体の指定がない run がないか
  検査4　用途ごとに文字の大きさが1つに揃っているか
  　　　 章14.0／節13.0／▌小見出し10.5／本文10.0／図9.0／出典8.5
  検査5　表内の文字の大きさが8.5pt（目次の表は10.0pt）に揃っているか
  検査6　本文が太字になっていないか（●要点・▌小見出し・章・節を除く）
  検査7　本文に強調記号（**）が残っていないか
  検査8　表紙・ヘッダー・本文に版の表記（Ver.／本版／第N版）が残っていないか
  　　　 （令和8年10月23日改め。素案は策定委員会に未提出のため版を紙面に出さず、
  　　　  版の番号はファイル名で管理する。旧の検査8は表紙とヘッダーの版の一致）
  検査9　行送りを指定した段落に lineRule があるか
  　　　 （ないと Word は w:line を倍率として読み、書体次第で行送りが変わる）
  検査10 表の列幅が実寸か（5mm未満の列・本文幅と食い違う表がないか）
  検査11 旧の書体（游ゴシック）の指定が残っていないか

使い方：
  python3 07_ソーススクリプト/check_font_R8.9.30.py \
      01_第10期_最新版成果品/川崎町_計画書素案_v2.34_協議用.docx
"""
import collections
import re
import sys
import zipfile

import docx
from docx.oxml.ns import qn

# 令和8年10月23日　第9期計画に揃えて HG丸ゴシックM-PRO とした（確認事項No.187）。
# Windows の登録名は半角カナ。全角で書くと Word が見つけられない
GOTHIC = "HG丸ｺﾞｼｯｸM-PRO"
KYU_GOTHIC = "游ゴシック"
RE_SHO = re.compile(r"^第\s*\d+\s*章$")
RE_SETSU = re.compile(r"^\d+-\d+　")
RE_CAP = re.compile(r"^図\d+-\d+　")
WANT = {"章": 14.0, "節": 13.0, "小見出し": 10.5, "本文": 10.0,
        "図": 9.0, "出典": 8.5}


def role_of(t):
    t = t.strip()
    if not t:
        return None
    if RE_SHO.match(t):
        return "章"
    if RE_SETSU.match(t):
        return "節"
    if t.startswith("▌") or t.startswith("　▌"):
        return "小見出し"
    if RE_CAP.match(t):
        return "図"
    if t.startswith("出典：") or t.startswith("資料："):
        return "出典"
    return "本文"


def sz(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return None
    e = rpr.find(qn("w:sz"))
    return int(e.get(qn("w:val"))) / 2 if e is not None else None


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else (
        "01_第10期_最新版成果品/川崎町_計画書素案_v2.34_協議用.docx")
    doc = docx.Document(path)
    body = doc.element.body
    kids = list(body.iterchildren())
    ok, ng = [], []

    def chk(name, cond, detail):
        (ok if cond else ng).append(f"{name}　{detail}")

    with zipfile.ZipFile(path) as z:
        st = z.read("word/styles.xml").decode("utf-8")
        hd = (z.read("word/header1.xml").decode("utf-8")
              if "word/header1.xml" in z.namelist() else "")

    # 検査1
    chk("検査1", 'w:styleId="Normal"' in st,
        "既定の段落スタイル Normal が" + ("ある" if 'w:styleId="Normal"' in st
                                    else "ない（Word が明朝体を当てるおそれ）"))

    # 検査2
    nofont = []
    for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>',
                         st, re.S):
        if GOTHIC not in m.group(0):
            nofont.append(m.group(1))
    chk("検査2", not nofont,
        f"{GOTHIC} の指定がないスタイル {len(nofont)}件"
        + ("：" + "／".join(nofont[:6]) if nofont else ""))

    # 検査3
    n_nofont = sum(1 for r in body.iter(qn("w:r"))
                   if r.find(qn("w:rPr")) is None
                   or r.find(qn("w:rPr")).find(qn("w:rFonts")) is None)
    chk("検査3", n_nofont == 0, f"書体の指定がない run {n_nofont}")

    # 目次の表の位置
    toc_idx = None
    for i, el in enumerate(kids):
        if el.tag == qn("w:tbl"):
            t = docx.table.Table(el, doc)
            txt = "".join(c.text for r in t.rows[:3] for c in r.cells)
            if "ご挨拶" in txt or "計画策定の背景と目的" in txt:
                toc_idx = i
                break
    if toc_idx is None:
        raise SystemExit("目次の表が見つからない")

    # 検査4
    by = collections.defaultdict(set)
    for i, el in enumerate(kids):
        if el.tag != qn("w:p") or i <= toc_idx:
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        role = role_of(t)
        if role is None:
            continue
        for r in el.iter(qn("w:r")):
            v = sz(r)
            if v:
                by[role].add(v)
    for role, want in WANT.items():
        got = sorted(by.get(role, set()))
        chk("検査4", got == [want] or not got,
            f"{role} の大きさ {got}（想定 {want}pt）")

    # 検査5
    tb = collections.Counter()
    for i, el in enumerate(kids):
        if el.tag != qn("w:tbl") or i == toc_idx:
            continue
        for r in el.iter(qn("w:r")):
            tb[sz(r)] += 1
    bad = {k: v for k, v in tb.items() if k not in (8.5, 10.5)}
    chk("検査5", not bad, f"表内の大きさ {dict(tb)}（想定 8.5pt・体系図のみ10.5pt）")

    # 検査6
    prev_sho = False
    bold = []
    for i, el in enumerate(kids):
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        is_sho = bool(RE_SHO.match(t))
        if i > toc_idx and t and not is_sho and not prev_sho:
            keep = (t.startswith("●") or t.startswith("▌")
                    or t.startswith("　▌") or RE_SETSU.match(t))
            if not keep:
                for r in el.iter(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is not None and rpr.find(qn("w:b")) is not None:
                        bold.append(t[:30])
                        break
        if t:
            prev_sho = is_sho
    chk("検査6", not bold,
        f"太字の本文 {len(bold)}段落" + ("：" + bold[0] if bold else ""))

    # 検査7
    ast = [p.text[:30] for p in doc.paragraphs if "**" in p.text]
    chk("検査7", not ast,
        f"強調記号（**）が残る段落 {len(ast)}" + ("：" + ast[0] if ast else ""))

    # 検査8
    han = re.compile(r"Ver\.\s?\d|本版|版です|第\d+版|範囲明記版")
    zen = "\n".join([p.text for p in doc.paragraphs]
                    + [c.text for t in doc.tables for r in t.rows
                       for c in r.cells]) + "\n" + re.sub(r"<[^>]+>", "", hd)
    hit = sorted({m.group(0) for m in han.finditer(zen)})
    chk("検査8", not hit, "版の表記 " + ("／".join(hit) if hit else "0件"))

    # 検査9
    nashi = 0
    for p in doc.element.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        sp = pPr.find(qn("w:spacing")) if pPr is not None else None
        if (sp is not None and sp.get(qn("w:line"))
                and not sp.get(qn("w:lineRule"))):
            nashi += 1
    chk("検査9", nashi == 0, f"lineRule のない行送り {nashi}段落")

    # 検査10
    sec = doc.sections[0]
    honbun = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    hen = []
    for ti, t in enumerate(doc.tables):
        ws = [int(g.get(qn("w:w")))
              for g in t._tbl.tblGrid.findall(qn("w:gridCol"))]
        if ws and (min(ws) < 283 or abs(sum(ws) - honbun) > honbun * 0.05):
            hen.append(ti)
    chk("検査10", not hen, f"列幅が実寸でない表 {len(hen)}件 {hen[:6]}")

    # 検査11
    with zipfile.ZipFile(path) as z:
        xml = "".join(z.read(n).decode("utf-8") for n in z.namelist()
                      if n.startswith("word/") and n.endswith(".xml"))
    chk("検査11", KYU_GOTHIC not in xml,
        f"旧の書体（{KYU_GOTHIC}）の指定 {xml.count(KYU_GOTHIC)}か所")

    print(f"書体の点検 {path}")
    for s in ng:
        print("  ✗", s)
    print(f"  適合 {len(ok)}件／不適合 {len(ng)}件")
    if ng:
        sys.exit(1)


if __name__ == "__main__":
    main()
