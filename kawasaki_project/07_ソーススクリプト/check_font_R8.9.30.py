# -*- coding: utf-8 -*-
"""書体と文字の大きさの点検

大雪地区広域連合の協議用素案（令和8年8月）の書式に合わせているかを確かめる。

  検査1　既定の段落スタイル（Normal）が定義されているか
  　　　 Title・Heading1〜6 はいずれも basedOn="Normal" を持つため、
  　　　 Normal がないと Word が自前の「標準」（明朝体のことがある）を当てる
  検査2　すべてのスタイルに游ゴシックが明示されているか
  検査3　書体の指定がない run がないか
  検査4　用途ごとに文字の大きさが1つに揃っているか
  　　　 章14.0／節13.0／▌小見出し10.5／本文10.0／図9.0／出典8.5
  検査5　表内の文字の大きさが8.5pt（目次の表は10.0pt）に揃っているか
  検査6　本文が太字になっていないか（●要点・▌小見出し・章・節を除く）
  検査7　本文に強調記号（**）が残っていないか
  検査8　ヘッダーの版表記が本文の版と一致しているか

使い方：
  python3 07_ソーススクリプト/check_font_R8.9.30.py \
      01_第10期_最新版成果品/川崎町_計画書素案_v2.9_事業内容版.docx
"""
import collections
import re
import sys
import zipfile

import docx
from docx.oxml.ns import qn

GOTHIC = "游ゴシック"
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
        "01_第10期_最新版成果品/川崎町_計画書素案_v2.9_事業内容版.docx")
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
        f"游ゴシックの指定がないスタイル {len(nofont)}件"
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
    m = re.search(r"計画書素案\s*(Ver\.[0-9.]+)", "\n".join(
        p.text for p in doc.paragraphs))
    mh = re.search(r"計画書素案\s*(Ver\.[0-9.]+)", re.sub(r"<[^>]+>", "", hd))
    chk("検査8", bool(m) and bool(mh) and m.group(1) == mh.group(1),
        f"本文の版 {m.group(1) if m else '不明'}／"
        f"ヘッダーの版 {mh.group(1) if mh else '不明'}")

    print(f"書体の点検 {path}")
    for s in ng:
        print("  ✗", s)
    print(f"  適合 {len(ok)}件／不適合 {len(ng)}件")
    if ng:
        sys.exit(1)


if __name__ == "__main__":
    main()
