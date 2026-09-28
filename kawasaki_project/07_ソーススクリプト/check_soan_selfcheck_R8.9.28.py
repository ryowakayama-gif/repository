# -*- coding: utf-8 -*-
"""計画素案の自己点検（交付金の数値の出典突合・同一事項の二値検査）

大雪広域連合の案件（ブランチ claude/大雪広域介護保険事業計画・f938242）で
外部レビューが指摘した3類型を、本案件でも機械で確かめられるようにしたもの。

  検査1　交付金の数値が出典（001732614.xlsx）から再現できるか
  検査2　同じ事項に2つの値が置かれていないか
  検査3　割合を書いた箇所に分母が併記されているか

不適合があれば終了コード1を返す。

使い方：
  python3 07_ソーススクリプト/check_soan_selfcheck_R8.9.28.py \
      01_第10期_最新版成果品/川崎町_計画書素案_v1.24_交付金枝番照合版.docx
"""
import json
import re
import sys

import docx
import openpyxl

XLSX = ("09_元資料/交付金評価/③令和８年度交付金評価指標等（市町村分・公表版）/"
        "001732614_令和８年度全国集計（市町村）.xlsx")
TAIKEI = ("05_試算・管理シート/"
          "川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx")
KAWASAKI = 280          # 全国一覧表の通し番号
COL = {"推進Ⅰ": 56, "推進Ⅱ": 78, "推進Ⅲ": 107, "推進Ⅳ": 140, "推進計": 141,
       "支援Ⅰ": 237, "支援Ⅱ": 266, "支援Ⅲ": 300, "支援Ⅳ": 333, "支援計": 334,
       "総合計": 335}


def load_source():
    ws = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)["全国集計（市町村）"]
    # 通し番号のない行（「標準偏差」等の集計行）を除く。
    # これを除かないと保険者数が1,742件・全国平均が454.9点になる。
    rows = [r for r in ws.iter_rows(min_row=15, max_row=1797, values_only=True)
            if r[0] is not None and isinstance(r[335], (int, float))]
    kawa = [r for r in rows if r[0] == KAWASAKI][0]
    # 指標群（体制・取組／活動／成果）の列は row4 の見出しから決める
    head = list(ws.iter_rows(min_row=4, max_row=4, values_only=True))[0] \
        if False else None
    return rows, kawa


def groups():
    """列番号の群分けを 00_交付金列定義.json から読む。無ければ作る。"""
    import os
    path = "04_調査・入力・分析/R8.9.25交付金再解析/交付金列定義_R8.9.28.json"
    if os.path.exists(path):
        return json.load(open(path, encoding="utf-8"))
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
    ws = wb["全国集計（市町村）"]
    r = {i: list(x) for i, x in
         enumerate(ws.iter_rows(min_row=1, max_row=14, values_only=True), 1)}

    def ffill(row):
        out, last = [], None
        for v in row:
            if v is not None and str(v).strip() != "":
                last = v
            out.append(last)
        return out

    SUB = ffill(r[4])
    EDA, EDA2, HAI = r[7], r[8], r[9]
    g = {"TAI": [], "KATSU": [], "SEI": []}
    for c in range(14, len(HAI)):
        if not isinstance(HAI[c], (int, float)):
            continue
        if not (EDA[c] or EDA2[c]):
            continue
        s = str(SUB[c] or "")
        if "体制" in s:
            g["TAI"].append(c)
        elif "活動" in s:
            g["KATSU"].append(c)
        elif "成果" in s:
            g["SEI"].append(c)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    json.dump(g, open(path, "w", encoding="utf-8"))
    return g


def text_of(doc):
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                out.append(cell.text)
    return "\n".join(out)


def main(path):
    rows, kawa = load_source()
    g = groups()

    def gs(r, k):
        return sum(r[i] for i in g[k] if isinstance(r[i], (int, float)))

    def rank(v, vals):
        return sorted(vals, reverse=True).index(v) + 1

    doc = docx.Document(path)
    body = text_of(doc)
    ng = []
    ok = 0

    # ══════════ 検査1　交付金の数値が出典から再現できるか
    want = []
    for k, c in COL.items():
        want.append((f"{k}の得点", str(kawa[c])))
    tot = [r[335] for r in rows]
    want.append(("総合計の全国順位", f"{rank(kawa[335], tot)}位"))
    want.append(("全国平均", f"{sum(tot) / len(tot):.1f}点"))
    want.append(("保険者数", f"{len(rows):,}保険者"))
    for name, key, man in (("体制・取組指標群", "TAI", 380),
                           ("活動指標群", "KATSU", 220),
                           ("成果指標群", "SEI", 200)):
        vals = [gs(r, key) for r in rows]
        mine = gs(kawa, key)
        want.append((f"{name}の得点", f"{mine}点"))
        want.append((f"{name}の全国順位", f"{rank(mine, vals):,}位"))
        want.append((f"{name}の全国平均", f"{sum(vals) / len(vals):.1f}点"))
    for name, val in want:
        if val.replace(",", "") in body.replace(",", ""):
            ok += 1
        else:
            ng.append(f"検査1　{name} の値「{val}」が素案に見当たらない")

    # ══════════ 検査2　同じ事項に2つの値が置かれていないか
    PAIR = [
        ("第1号被保険者負担割合24％の影響（月額）",
         re.compile(r"24％と(?:定められた場合|なった場合)は?、?月額約(\d+)円")),
        ("調整交付金5.0％の影響（月額）",
         re.compile(r"標準の5\.0％と?(?:なった場合)?は?、?月額約(\d+)円")),
        ("推進 目標Ⅲ の全国平均",
         re.compile(r"満点100点・全国平均(\d+\.\d+)点")),
        ("体制・取組指標群の失点の合計",
         re.compile(r"失点は(\d+)か所・(\d+)点")),
        ("第1号被保険者数",
         re.compile(r"第1号被保険者数(?:は)?(?:約)?([\d,]+)人")),
    ]
    for name, pat in PAIR:
        found = {m.group(0): m.groups() for m in pat.finditer(body)}
        vals = {v for v in found.values()}
        if len(vals) > 1:
            ng.append(f"検査2　{name} に2つ以上の値がある: "
                      + " / ".join(sorted(found)))
        elif vals:
            ok += 1

    # ══════════ 検査3　調査の割合に分母が併記されているか
    #   「○○.○％」の直後・直前30字以内に n= 又は「件」「人」があることを求める
    miss = []
    for p in doc.paragraphs:
        t = p.text
        if "調査" not in t and "ニーズ" not in t and "実態" not in t:
            continue
        for m in re.finditer(r"(\d+\.\d)％", t):
            lo = max(0, m.start() - 40)
            ctx = t[lo:m.end() + 40]
            if re.search(r"n[=＝]\d|\d+件|\d+人|ポイント|全国|県", ctx):
                continue
            miss.append(f"{t[:14]}… {m.group(0)}")
    if miss:
        ng.append(f"検査3　分母の併記がない調査の割合 {len(miss)}件: "
                  + " / ".join(miss[:5]))
    else:
        ok += 1

    # ══════════ 検査4　柱別の事業数の内訳が、素案の掲げる合計と一致するか
    #   第4章の表42は柱ごとに「（n節m項k事業）」を記している。
    #   k の合計が、冒頭・第4章本文が掲げる総事業数と一致しなければならない。
    # 柱の一覧表（見出しが「柱／内容」）の中だけを見る。
    # 第5章・第6章の節の見出しにも「（3項9事業）」の形が現れるため、
    # 本文全体を対象にすると二重に数えてしまう。
    hashira = ""
    for t in doc.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if h[:2] == ["柱", "内容"]:
            hashira = "\n".join(c.text for r in t.rows for c in r.cells)
    uchiwake = [int(m.group(3)) for m in
                re.finditer(r"（(\d+)節(\d+)項(\d+)事業）", hashira)]
    uchiwake += [int(m.group(2)) for m in
                 re.finditer(r"（(\d+)項(\d+)事業）", hashira)]
    goukei = {int(m.group(3)) for m in
              re.finditer(r"全7章(\d+)節(\d+)項(\d+)事業", body)}
    if uchiwake and goukei:
        if len(goukei) > 1:
            ng.append(f"検査4　総事業数に2つ以上の値がある: {sorted(goukei)}")
        elif sum(uchiwake) != list(goukei)[0]:
            ng.append(f"検査4　柱別の事業数の合計 {sum(uchiwake)} が"
                      f"、掲げている総事業数 {list(goukei)[0]} と一致しない")
        else:
            ok += 1
    elif goukei and not hashira:
        ng.append("検査4　柱の一覧表（柱／内容）が見つからない")

    # ══════════ 検査5　書き下ろした事業が統合体系表と一致するか
    #   Ver.2.0で第5章・第6章に置いた事業の表（項／事業／区分／所管）と、
    #   第8章の柱7の対応表を、統合体系表の全147行（143事業＋再掲4）と突き合わせる。
    import os                                     # noqa: PLC0415
    from collections import Counter               # noqa: PLC0415
    if os.path.exists(TAIKEI):
        tw = openpyxl.load_workbook(TAIKEI, data_only=True,
                                    read_only=True)["01_統合体系表"]
        src = Counter()
        n7 = 0
        for r in tw.iter_rows(min_row=2, values_only=True):
            if not r[0]:
                continue
            if str(r[0]).startswith("第７章"):
                n7 += 1
                continue
            src[str(r[3]).strip()] += 1
        got = Counter()
        n7doc = 0
        for t in doc.tables:
            h = [c.text.strip() for c in t.rows[0].cells]
            if h[:4] == ["項", "事業", "区分", "所管"]:
                for row in t.rows[1:]:
                    got[row.cells[1].text.strip()] += 1
            elif h[:3] == ["柱7の事業", "区分", "本計画での位置"]:
                n7doc = len(t.rows) - 1
        if sum(got.values()) == 0:
            pass          # Ver.1.x は書き下ろし前なので検査しない
        else:
            miss = src - got
            extra = got - src
            if miss or extra:
                ng.append(f"検査5　書き下ろした事業が統合体系表と一致しない"
                          f"（不足{sum(miss.values())}件・余分{sum(extra.values())}件）"
                          + (" 不足: " + "／".join(list(miss)[:3]) if miss else "")
                          + (" 余分: " + "／".join(list(extra)[:3]) if extra else ""))
            else:
                ok += 1
            if n7doc != n7:
                ng.append(f"検査5　柱7の事業が{n7doc}件（統合体系表は{n7}件）")
            else:
                ok += 1

    print(f"自己点検 {path}")
    print(f"  適合 {ok}件／不適合 {len(ng)}件")
    for m in ng:
        print("  ×", m)
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else
                  "01_第10期_最新版成果品/"
                  "川崎町_計画書素案_v1.24_交付金枝番照合版.docx"))
