# -*- coding: utf-8 -*-
"""章節の参照の機械検索（令和8年10月24日 Red Team レビュー A12 への対応）

  python3 07_ソーススクリプト/check_sansho_R8.10.24.py [素案の docx]

素案の本文・表に現れる参照を拾い、参照先が素案に実在するかを確かめる。
  ・「N-M」（節）……素案に「N-M　」で始まる見出しがあるか
  ・「第N章」……素案に「第 N 章」の章扉があるか
  ・「第N章N-M」「第N章のN-M」……章番号と節番号が食い違っていないか（例「第7章9-3」）
加えて、旧の誤り「詳細は第7章で論じます」（所得段階）が残っていないかを見る。

参照の**内容**（その節に参照どおりのことが書いてあるか）は機械では確かめられない。
章を名指しする文（「第N章で」）は一覧に出し、人が確かめる（未実施を適合に丸めない）。

終了コード　0＝実在しない参照・章と節の食い違い 0件／1＝あり
"""
import collections
import re
import sys

import docx
from docx.oxml.ns import qn

SOAN = sys.argv[1] if len(sys.argv) > 1 else "01_第10期_最新版成果品/川崎町_計画書素案_v2.36_協議用.docx"
ZEN = str.maketrans("０１２３４５６７８９", "0123456789")


def main():
    d = docx.Document(SOAN)
    body = d.element.body
    sho, setsu = set(), set()
    bun = []           # (節, 文)
    cur = None
    for el in body.iterchildren():
        if el.tag == qn("w:tbl"):
            # 表はセルの段落ごとに区切って読む（つなげると「8-1」「2」が「8-12」に見える）
            t = "\n".join("".join(n.text or "" for n in q.iter(qn("w:t")))
                           for q in el.iter(qn("w:p"))).translate(ZEN)
        else:
            t = "".join(n.text or "" for n in el.iter(qn("w:t"))).translate(ZEN)
        if el.tag == qn("w:p"):
            s = t.strip()
            m = re.match(r"^第\s*(\d+)\s*章$", s)
            if m:
                sho.add(int(m.group(1)))
            m = re.match(r"^(\d+)-(\d+)[　 ]", s)
            if m:
                setsu.add((int(m.group(1)), int(m.group(2))))
                cur = f"{m.group(1)}-{m.group(2)}"
                continue
        bun.append((cur, t))
    print(f"章節の参照の機械検索　{SOAN}")
    print(f"  章 {len(sho)}件（{min(sho)}〜{max(sho)}）・節 {len(setsu)}件")
    nai, kuichigai, shomei = [], [], []
    n_setsu = n_sho = 0
    for sec, t in bun:
        # 目次・図表番号一覧の行（「1-1．…」）は参照ではない
        for m in re.finditer(r"(?<![\d.,／/\-－])(\d{1,2})-(\d{1,2})(?![\d.,％%年月日人件円])", t):
            a, b = int(m.group(1)), int(m.group(2))
            if a > 10 or b > 12:
                continue
            n_setsu += 1
            if (a, b) not in setsu:
                nai.append((sec, f"{a}-{b}", t[max(0, m.start() - 30):m.end() + 20]))
        for m in re.finditer(r"第(\d{1,2})章[のの]?(\d{1,2})-(\d{1,2})", t):
            if int(m.group(1)) != int(m.group(2)):
                kuichigai.append((sec, m.group(0), t[max(0, m.start() - 30):m.end() + 20]))
        for m in re.finditer(r"第(\d{1,2})章", t):
            n_sho += 1
            if int(m.group(1)) not in sho:
                nai.append((sec, m.group(0), t[max(0, m.start() - 30):m.end() + 20]))
            if re.match(r"第\d{1,2}章(で|に|を|の各|参照)", t[m.start():m.end() + 3]):
                shomei.append((sec, m.group(0), t[max(0, m.start() - 40):m.end() + 25]))
    print(f"  拾った参照　節 {n_setsu}件・章 {n_sho}件")
    print(f"  ── 検査1　実在しない参照先　{len(nai)}件")
    for s, r, c in nai[:30]:
        print(f"   × [{s}] {r}　…{c}…")
    print(f"  ── 検査2　章と節の番号の食い違い　{len(kuichigai)}件")
    for s, r, c in kuichigai:
        print(f"   × [{s}] {r}　…{c}…")
    ato = "\n".join(t for _, t in bun)
    kyu = "詳細は第7章で論じます" in ato
    print(f"  ── 検査3　旧の誤り「詳細は第7章で論じます」　{'残っている' if kyu else '0件'}")
    print(f"  ── 参考　章を名指しする文 {len(shomei)}件（内容の当否は人が確かめる）")
    c = collections.Counter(r for _, r, _ in shomei)
    print("    " + "／".join(f"{k} {v}件" for k, v in sorted(c.items(), key=lambda x: int(re.sub(r'\D', '', x[0])))))
    for s, r, t in shomei:
        print(f"    [{s}] …{t}…")
    ng = len(nai) + len(kuichigai) + int(kyu)
    print(("   ○ " if not ng else "   × ") + f"不適合 {ng}件")
    sys.exit(1 if ng else 0)


if __name__ == "__main__":
    main()
