# -*- coding: utf-8 -*-
"""資料編 付表1 と 本編（介護実態調査結果報告書）の同一設問の値を直接突き合わせる。

個票を経由せず、2つの文書に書かれた数字そのものを比べる。
本編と資料編を並べて読む委員・監査の目線に一番近い照合。

あわせて、単一回答設問の割合の合計が 100.0% になっているかを確認する
（端数処理の説明が3文書とも欠落しているため、合計が100.0にならない表を洗い出す）。

使い方
  python3 verify_shiryo_vs_honpen.py [出力txt]
"""
import re
import sys
import unicodedata

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = "/home/user/repository/kawasaki_project/09_元資料/R8調査データ/R8.9.9受領版"
SHIRYO = f"{BASE}/川崎町_資料編_R8.9.9版.docx"
HONPEN = f"{BASE}/川崎町_介護実態調査結果報告書_R8.9.9版.docx"

# 付表1 → 本編の対応する表（単純集計）
MAP = {
    "付表1-1": "表3-5", "付表1-2": "表3-2", "付表1-3": "表3-3",
    "付表1-4": "表3-4", "付表1-5": "表4-2", "付表1-6": "表4-3",
    "付表1-7": "表5-1", "付表1-20": "表6-1", "付表1-21": "表8-2",
    "付表1-22": "表7-1", "付表1-23": "表7-2", "付表1-24": "表8-1",
    "付表1-25": "表8-3", "付表1-26": "表8-4",
}
# 付表1-8〜1-19（A問8）はサービス別。本編では 表5-3（利用の有無）と
# 表5-4（利用頻度）に横持ちで収録されている
SVC_ORDER = [
    "訪問介護（ホームヘルプ）", "訪問入浴介護", "訪問看護", "訪問リハビリテーション",
    "通所介護（デイサービス）", "通所リハビリテーション（デイケア）",
    "夜間対応型訪問介護", "定期巡回・随時対応型訪問介護看護",
    "小規模多機能型居宅介護", "看護小規模多機能型居宅介護",
    "短期入所生活介護・療養介護（ショートステイ）", "居宅療養管理指導",
]

# 単一回答の設問（割合の合計が100.0になるべきもの）
SINGLE = {"付表1-2", "付表1-3", "付表1-4", "付表1-5", "付表1-6", "付表1-7",
          "付表1-8", "付表1-9", "付表1-10", "付表1-11", "付表1-12", "付表1-13",
          "付表1-14", "付表1-15", "付表1-16", "付表1-17", "付表1-18", "付表1-19",
          "付表1-20", "付表1-22", "付表1-24", "付表1-26"}


def norm(s):
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[\s　]+", "", s)
    s = re.sub(r"[（）()［］\[\]「」・,、。／/？?]", "", s)
    return s.replace("～", "-").replace("〜", "-").replace("~", "-")


def blocks(d):
    for ch in d.element.body.iterchildren():
        if ch.tag == qn("w:p"):
            yield Paragraph(ch, d)
        elif ch.tag == qn("w:tbl"):
            yield Table(ch, d)


def ctext(c):
    return "\n".join(p.text.strip() for p in c.paragraphs if p.text.strip())


def read_tables(path):
    """キャプション（先頭の表番号）→ グリッド"""
    d = Document(path)
    out, cap = {}, ""
    for b in blocks(d):
        if isinstance(b, Paragraph):
            if b.text.strip():
                cap = b.text.strip()
            continue
        tag = re.match(r"^(付表[12]-\d+|表\d+-\d+)", cap)
        if tag:
            out[tag.group(1)] = ([[ctext(c) for c in r.cells] for r in b.rows], cap)
    return out


def num(s):
    """セルから 件数 と 割合 を取り出す"""
    s = s.replace(" ", "").replace(",", "").replace("件", "")
    k = re.match(r"^([0-9]+)", s)
    p = re.search(r"([0-9]+\.[0-9]+)\s*[%％]", s)
    return (int(k.group(1)) if k else None,
            float(p.group(1)) if p else None)


S = read_tables(SHIRYO)
H = read_tables(HONPEN)

L = ["川崎町 資料編（R8.9.9版）×本編 の突合",
     f"資料編：{SHIRYO.split('/')[-1]}",
     f"本編　：{HONPEN.split('/')[-1]}", ""]

cmp_n = dif_n = miss = 0
difs, misses = [], []

# ---------------------------------------------- 単純集計（1対1で対応する14表）
for sk, hk in MAP.items():
    if sk not in S:
        misses.append(f"{sk} が資料編にない")
        continue
    if hk not in H:
        misses.append(f"{hk} が本編にない（{sk} の対照先）")
        continue
    sg, scap = S[sk]
    hg, hcap = H[hk]
    hmap = {}
    for r in hg[1:]:
        if len(r) >= 2:
            hmap[norm(r[0])] = r[1:]
    for r in sg[1:]:
        lab = norm(r[0])
        if lab in ("有効回答数n", "有効回答数", "計", "合計", "n", "総数"):
            continue
        if lab not in hmap:
            miss += 1
            misses.append(f"{sk}「{r[0][:20]}」に対応する行が {hk} にない")
            continue
        sk_, sp_ = None, None
        for cell in r[1:]:
            a, b = num(cell)
            if a is not None and sk_ is None and b is None:
                sk_ = a
            if b is not None:
                sp_ = b
                if sk_ is None:
                    sk_ = a
        hk_, hp_ = None, None
        for cell in hmap[lab]:
            a, b = num(cell)
            if a is not None and hk_ is None and b is None:
                hk_ = a
            if b is not None:
                hp_ = b
                if hk_ is None:
                    hk_ = a
        for what, sv, hv in (("件数", sk_, hk_), ("割合", sp_, hp_)):
            if sv is None or hv is None:
                continue
            cmp_n += 1
            if abs(sv - hv) > 0.001:
                dif_n += 1
                difs.append(f"{sk}／{hk}「{r[0][:20]}」{what} 資料編{sv} ≠ 本編{hv}")

# ---------------------------------------------- A問8（サービス別）：表5-3 と対照
if "表5-3" in H:
    hg, _ = H["表5-3"]
    head = [norm(x) for x in hg[0]]
    for i, svc in enumerate(SVC_ORDER):
        sk = f"付表1-{8+i}"
        if sk not in S:
            misses.append(f"{sk}（{svc}）が資料編にない")
            continue
        sg, _ = S[sk]
        # 本編 表5-3 の該当行
        hrow = next((r for r in hg[1:] if norm(r[0]) == norm(svc)), None)
        if hrow is None:
            miss += 1
            misses.append(f"表5-3 に「{svc}」の行がない（{sk} の対照先）")
            continue
        # 資料編側：「利用していない」以外の合計＝利用あり
        s_no = s_yes = 0
        for r in sg[1:]:
            lab = norm(r[0])
            if lab in ("有効回答数n", "有効回答数", "計", "合計", "n", "総数"):
                continue
            k, _ = num(r[1]) if len(r) > 1 else (None, None)
            if k is None:
                continue
            if lab == norm("利用していない"):
                s_no += k
            else:
                s_yes += k
        # 本編側：「利用した」に当たる列を探す
        h_yes = None
        for j, h in enumerate(head):
            if h in (norm("利用した"), norm("利用あり"), norm("利用している")):
                h_yes = num(hrow[j])[0]
        if h_yes is None:
            misses.append(f"表5-3 に「利用した」列が見当たらない（{sk}）")
            continue
        cmp_n += 1
        if s_yes != h_yes:
            dif_n += 1
            difs.append(f"{sk}／表5-3「{svc[:18]}」利用あり 資料編{s_yes} ≠ 本編{h_yes}")

L.append(f"■ 資料編 付表1 と 本編の同一設問の突合")
L.append(f"  照合 {cmp_n} 項目／相違 {dif_n} 件／対照先が見つからなかった行 {miss} 件")
for x in difs:
    L.append("  ✗ " + x)
if not difs:
    L.append("  （値の相違なし）")
for x in misses[:25]:
    L.append("  ? " + x)
if len(misses) > 25:
    L.append(f"  …ほか {len(misses)-25} 件")
L.append("")

# ---------------------------------------------- 割合の合計が100.0か
L.append("■ 単一回答設問の割合の合計（端数処理の確認）")
ng = 0
for sk in sorted(SINGLE, key=lambda x: int(x.split("-")[1])):
    if sk not in S:
        continue
    sg, scap = S[sk]
    tot = 0.0
    for r in sg[1:]:
        lab = norm(r[0])
        if lab in ("有効回答数n", "有効回答数", "計", "合計", "n", "総数"):
            continue
        for cell in r[1:]:
            _, p = num(cell)
            if p is not None:
                tot += p
                break
    if abs(tot - 100.0) > 0.001:
        ng += 1
        L.append(f"  △ {sk} 合計 {round(tot,1)}%　{scap[:44]}")
L.append(f"  合計が100.0%にならない表：{ng} / {len(SINGLE)} 表"
         "　→ 端数処理の注記が必要")
L.append("")

out = sys.argv[1] if len(sys.argv) > 1 else "chk_資料編×本編.txt"
with open(out, "w", encoding="utf-8") as f:
    f.write("\n".join(L) + "\n")
print("\n".join(L))
print(f"→ {out}")
