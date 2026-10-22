# -*- coding: utf-8 -*-
"""保険者機能強化推進交付金等の推移の再整理（令和8年10月22日）

  python3 07_ソーススクリプト/analyze_kofukin_suii_R8.10.22.py

なぜ要るか
  同じ事務所の他の案件（2件）が、同じ原典を別の順序で読み、
  交付金の整理で次の誤りを見つけていた。本町でも起きていないかを、
  **原典から数え直して**確かめる。他団体の数値は用いない。

    ① 枝の配点の見かけの合計（929点）を満点と取り違える
       （選択肢が排他の枝を含むため、取り得るのは800点）
    ② 成果指標群（目標Ⅳ）は推進・支援で同一の指標であり、
       **項目の数を2度数える**（点は両交付金に別に入るので足してよい）
    ③ 推移を**素点**で比べると、国が配点を下げただけの項目が
       「悪化」に入る。**得点率**で比べる
    ④ 前年度からの低下を、取組の後退・成果指標の変動・配点の改定に
       分けずに示す
    ⑤ 全国平均が毎年上がっていることを示さずに順位の低下を示す
    ⑥ 県内の位置（団体名を出さずに分布で示す）
    ⑦ 交付額（円）を把握しているか

  ⚠ **他団体の資料は「問いの立て方」として用いた。数は借りていない。**
    県内の比較は分布（平均・四分位・中央値）と本町の順位だけを用い、
    他団体の名称・個別の得点は書き出さない。

入力　09_元資料/交付金評価/①令和６〜８年度…/01_（掲載用）…評価結果（市町村分）….xlsx
出すもの　05_試算・管理シート/川崎町_交付金_推移の再整理_R8.10.22.xlsx

終了コード　0＝検査1〜9すべて適合／1＝1件でも不適合
"""
import glob
import re
import statistics
import sys
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NENDO = {6: "６", 7: "７", 8: "８"}
OUT = "05_試算・管理シート/川崎町_交付金_推移の再整理_R8.10.22.xlsx"
KEN = 4                     # 宮城県の都道府県番号
MACHI = "川崎町"
# 以前に本町の資料に書いた値（再現できるかを確かめる）
KAKOU = {6: 454, 7: 565, 8: 528}

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="D9E1F2")
WARN = PatternFill("solid", fgColor="FCE4D6")
OK = PatternFill("solid", fgColor="E2EFDA")
WRAP = Alignment(vertical="top", wrap_text=True)


def s(v):
    return "" if v is None else re.sub(r"\s+", "", str(v))


def ffill(ws, r, n):
    out, last = [], ""
    for c in range(1, n + 1):
        v = s(ws.cell(r, c).value)
        if v:
            last = v
        out.append(last)
    return out


def yomu(y):
    p = glob.glob(f"09_元資料/交付金評価/①令和{NENDO[y]}年度*/01_*.xlsx")[0]
    ws = openpyxl.load_workbook(p, data_only=True).worksheets[0]
    n = ws.max_column
    raw = {r: [s(ws.cell(r, c).value) for c in range(1, n + 1)]
           for r in range(2, 9)}
    kofu, moku = ffill(ws, 2, n), ffill(ws, 3, n)
    gun, bango, na = ffill(ws, 4, n), ffill(ws, 5, n), ffill(ws, 6, n)
    # ⚠ 枝は7行目（ア・イ…）と8行目（①・②…）の2段。8行目を落とすと
    #   同じ鍵が重なり、上書きされて点が消える（現にそれで合計が合わなかった）
    eda7 = ffill(ws, 7, n)
    eda = [(raw[7][c] or eda7[c]) + raw[8][c] if raw[8][c] else raw[7][c]
           for c in range(n)]
    hai = [ws.cell(9, c).value for c in range(1, n + 1)]

    # 交付金の範囲（最初の「保険者機能強化推進交付金」から「推進・支援合計」まで）
    c0 = next(c for c in range(n) if raw[2][c].startswith("保険者機能強化推進"))
    c_tot = next(c for c in range(c0, n)
                 if "推進・支援合計" in raw[2][c] or "推進・支援合計" in raw[3][c])
    # 葉の列（枝そのもの）＝配点が数で、見出しが集計（〜計・合計）でないもの
    leaf = []
    for c in range(c0, c_tot):
        if not isinstance(hai[c], (int, float)):
            continue
        agg = any(raw[r][c].endswith("計") or "合計" in raw[r][c]
                  for r in (2, 3, 4, 5, 6))
        if agg:
            continue
        k = "推進" if "推進" in kofu[c] else "支援"
        mm = re.search(r"目標([ⅠⅡⅢⅣ])", moku[c])
        mk = mm.group(1) if mm else ""
        g = ("ⅰ" if "（ⅰ）" in gun[c] else "ⅱ" if "（ⅱ）" in gun[c]
             else "成果" if (mk == "Ⅳ" or "成果" in gun[c]) else gun[c][:6])
        leaf.append(dict(col=c + 1, kofu=k, moku=mk, gun=g, no=bango[c],
                         na=na[c], eda=eda[c], hai=float(hai[c])))
    # 同じ鍵が残るときは出現の順で区別する（年度をまたいでも順は保たれる）
    seen = {}
    for l in leaf:
        k0 = (l["kofu"], l["moku"], l["gun"], l["na"], l["eda"])
        seen[k0] = seen.get(k0, 0) + 1
        l["jun"] = seen[k0]

    # 保険者の行（都道府県番号が数のもの）
    rows = [r for r in range(10, ws.max_row + 1)
            if isinstance(ws.cell(r, 2).value, (int, float))
            and s(ws.cell(r, 7).value or ws.cell(r, 8).value)
            and isinstance(ws.cell(r, c_tot + 1).value, (int, float))]
    me = next(r for r in rows if ws.cell(r, 2).value == KEN and
              MACHI in (s(ws.cell(r, 7).value), s(ws.cell(r, 8).value)))
    tot = {r: ws.cell(r, c_tot + 1).value for r in rows}
    return dict(ws=ws, leaf=leaf, rows=rows, me=me, tot=tot,
                ken=[r for r in rows if ws.cell(r, 2).value == KEN])


def key(l):
    return (l["kofu"], l["moku"], l["gun"], l["na"], l["eda"], l["jun"])


def sheet(wb, name, head, widths):
    ws = wb.create_sheet(name)
    for c, (h, w) in enumerate(zip(head, widths), start=1):
        x = ws.cell(1, c, h)
        x.font = Font(bold=True)
        x.fill = HEAD
        x.border = BOX
        x.alignment = WRAP
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    return ws


def put(ws, rows, fill=None):
    for i, r in enumerate(rows, start=2):
        for j, v in enumerate(r, start=1):
            x = ws.cell(i, j, v)
            x.border = BOX
            x.alignment = WRAP
        f = fill(r) if fill else None
        if f:
            for j in range(1, len(r) + 1):
                ws.cell(i, j).fill = f


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    Y = {y: yomu(y) for y in (6, 7, 8)}
    print("交付金の推移の再整理（令和6〜8年度・原典から）")
    print()

    # ══════ 検査1　合計が再現できるか（以前に書いた値と一致するか）
    print("  ── 検査1　本町の合計得点（葉の列の和と原典の合計欄）")
    me_leaf = {}
    for y, d in Y.items():
        ws, me = d["ws"], d["me"]
        v = {key(l): (ws.cell(me, l["col"]).value or 0, l["hai"])
             for l in d["leaf"]}
        me_leaf[y] = v
        judge(len(v) == len(d["leaf"]), f"令和{y}年度 鍵の重なり0"
              f"（枝{len(d['leaf'])}・鍵{len(v)}）")
        wa = sum(a for a, _ in v.values())
        judge(abs(wa - d["tot"][me]) < 0.01 and d["tot"][me] == KAKOU[y],
              f"令和{y}年度 {d['tot'][me]:.0f}点（葉の和 {wa:.0f}・"
              f"以前の記載 {KAKOU[y]}）")

    # ══════ 検査2　⚠ 見かけの配点と満点
    print("  ── 検査2　⚠ 枝の配点の見かけの合計と満点800点")
    for y, d in Y.items():
        mie = sum(l["hai"] for l in d["leaf"])
        print(f"     令和{y}年度　枝 {len(d['leaf'])}件・配点の見かけの合計 "
              f"{mie:.0f}点（⚠ 満点は800点）")
    judge(True, "満点は800点として扱う（失点＝800−得点）")

    # ══════ 全国・県内の位置
    print("  ── 全国・県内の位置")
    ichi = {}
    for y, d in Y.items():
        t = [d["tot"][r] for r in d["rows"]]
        me = d["tot"][d["me"]]
        k = [d["tot"][r] for r in d["ken"]]
        q = statistics.quantiles(k, n=4)
        ichi[y] = dict(
            me=me, n=len(t), avg=sum(t) / len(t),
            rank=sum(1 for v in t if v > me) + 1,
            kn=len(k), kavg=sum(k) / len(k), kmed=statistics.median(k),
            kq1=q[0], kq3=q[2], kmax=max(k), kmin=min(k),
            krank=sum(1 for v in k if v > me) + 1)
        i = ichi[y]
        print(f"     令和{y}年度　本町 {me:.0f}点　全国平均 {i['avg']:.1f}・"
              f"{i['rank']}位／{i['n']}　宮城県平均 {i['kavg']:.1f}・"
              f"中央値 {i['kmed']:.1f}・県内 {i['krank']}位／{i['kn']}")
    judge(ichi[8]["n"] == 1741, f"令和8年度の保険者数1,741（実際 {ichi[8]['n']}）")
    judge(ichi[8]["rank"] == 366, f"令和8年度 全国366位（実際 {ichi[8]['rank']}）")

    # ══════ 検査3　令和7→8年度の低下の内訳（枝ごと）
    print("  ── 検査3　令和7→8年度の増減を枝ごとに分ける")
    k7, k8 = me_leaf[7], me_leaf[8]
    bunrui = OrderedDict((b, []) for b in (
        "成果指標群の変動", "配点の改定（得点率は不変）", "指標の新設",
        "指標の廃止・組替え", "取組の後退（得点率の低下）",
        "取組の前進（得点率の上昇）"))
    for kk in sorted(set(k7) | set(k8)):
        a7, h7 = k7.get(kk, (None, None))
        a8, h8 = k8.get(kk, (None, None))
        d78 = (a8 or 0) - (a7 or 0)
        if d78 == 0:
            continue
        if kk[1] == "Ⅳ":
            b = "成果指標群の変動"
        elif a7 is None:
            b = "指標の新設"
        elif a8 is None:
            b = "指標の廃止・組替え"
        elif h7 and h8 and abs(a7 / h7 - a8 / h8) < 1e-9:
            b = "配点の改定（得点率は不変）"
        elif h7 and h8 and a8 / h8 < a7 / h7:
            b = "取組の後退（得点率の低下）"
        else:
            b = "取組の前進（得点率の上昇）"
        bunrui[b].append((kk, a7, h7, a8, h8, d78))
    gokei = 0
    for b, lst in bunrui.items():
        v = sum(x[5] for x in lst)
        gokei += v
        print(f"     {b:<22}{len(lst):>3}枝　{v:+.0f}点")
    d_tot = Y[8]["tot"][Y[8]["me"]] - Y[7]["tot"][Y[7]["me"]]
    judge(abs(gokei - d_tot) < 0.01,
          f"枝ごとの増減の計（{gokei:+.0f}）が合計の増減（{d_tot:+.0f}）と一致")
    kyotsu = len(set(k7) & set(k8))
    print(f"     （令和7・8年度で同じ枝と突き合わせられた数 {kyotsu}／"
          f"令和8年度 {len(k8)}枝）")

    # ══════ 検査4　⚠ 素点と得点率で「悪化」の数が変わるか（項目の単位）
    print("  ── 検査4　⚠ 素点で比べた「悪化」と得点率で比べた「悪化」")
    def koumoku(v):
        o = {}
        for kk, (a, h) in v.items():
            ik = kk[:4]
            e = o.setdefault(ik, [0, 0])
            e[0] += a
            e[1] += h
        return o
    i6, i8 = koumoku(me_leaf[6]), koumoku(me_leaf[8])
    sotten, ritsu, haiten_dake = [], [], []
    for ik in set(i6) & set(i8):
        a6, h6 = i6[ik]
        a8, h8 = i8[ik]
        if a8 < a6:
            sotten.append(ik)
            if h6 and h8 and a8 / h8 >= a6 / h6 - 1e-9:
                haiten_dake.append((ik, a6, h6, a8, h8))
        if h6 and h8 and a8 / h8 < a6 / h6 - 1e-9:
            ritsu.append(ik)
    print(f"     令和6→8年度で共通の項目 {len(set(i6) & set(i8))}件")
    print(f"     素点で「悪化」{len(sotten)}件／得点率で「悪化」{len(ritsu)}件")
    print(f"     ⚠ 素点では悪化・得点率では悪化していない（配点の改定だけ）"
          f" {len(haiten_dake)}件")
    for ik, a6, h6, a8, h8 in haiten_dake:
        print(f"        {ik[0]}{ik[1]}({ik[2]}) {ik[3][:24]}　"
              f"{a6:.0f}/{h6:.0f} → {a8:.0f}/{h8:.0f}")

    # ══════ 検査5　⚠ 目標Ⅳの二重
    print("  ── 検査5　⚠ 成果指標群（目標Ⅳ）は推進・支援で同一か")
    s4 = {kk[3:]: v for kk, v in k8.items() if kk[0] == "推進" and kk[1] == "Ⅳ"}
    t4 = {kk[3:]: v for kk, v in k8.items() if kk[0] == "支援" and kk[1] == "Ⅳ"}
    judge(set(s4) == set(t4) and all(s4[k] == t4[k] for k in s4),
          f"推進Ⅳ{len(s4)}枝と支援Ⅳ{len(t4)}枝の指標・配点・得点が一致",
          "（点は両交付金に別に入るので足してよい。項目の数は1度だけ数える）")

    # ══════ 検査6　本町の強みとして書いた値を再現できるか
    print("  ── 検査6　委員会資料「本町の強み」の値")
    d8 = Y[8]

    def moku_tot(kofu, mk):
        cols = [l["col"] for l in d8["leaf"]
                if l["kofu"] == kofu and l["moku"] == mk]
        val = {r: sum(d8["ws"].cell(r, c).value or 0 for c in cols)
               for r in d8["rows"]}
        me = val[d8["me"]]
        return (me, sum(1 for v in val.values() if v > me) + 1,
                sum(val.values()) / len(val),
                sum(1 for r in d8["ken"] if val[r] > me) + 1)
    me1, rk1, av1, kr1 = moku_tot("支援", "Ⅰ")
    judge(me1 == 87 and rk1 == 5,
          f"支援 目標Ⅰ 87点・全国5位（実際 {me1:.0f}点・{rk1}位・"
          f"全国平均{av1:.1f}・県内{kr1}位）")
    me3, rk3, av3, kr3 = moku_tot("推進", "Ⅲ")
    judge(me3 == 73 and rk3 == 218,
          f"推進 目標Ⅲ 73点・全国218位（実際 {me3:.0f}点・{rk3}位）")
    # 難聴（推進・支援のどこにあるかを名で探す）
    nan = [l for l in d8["leaf"] if "難聴" in l["na"]]
    if nan:
        cols = [l["col"] for l in nan]
        val = {r: sum(d8["ws"].cell(r, c).value or 0 for c in cols)
               for r in d8["rows"]}
        me = val[d8["me"]]
        kv = [val[r] for r in d8["ken"]]
        kr = sum(1 for v in kv if v > me) + 1
        judge(me == sum(l["hai"] for l in nan) and kr == 1,
              f"難聴 {me:.0f}/{sum(l['hai'] for l in nan):.0f}点・県内{kr}位"
              f"（県平均{sum(kv) / len(kv):.1f}・"
              f"全国平均{sum(val.values()) / len(val):.1f}）")
    else:
        judge(False, "難聴の枝が見つからない")

    # ══════ 検査7　目標ごとの県内の位置（団体名を出さない）
    print("  ── 検査7　目標ごとの全国・県内の位置（令和8年度）")
    mok = []
    for kofu in ("推進", "支援"):
        for mk in ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"):
            if kofu == "支援" and mk == "Ⅳ":
                continue        # 推進Ⅳと同一
            me, rk, av, kr = moku_tot(kofu, mk)
            cols = [l["col"] for l in d8["leaf"]
                    if l["kofu"] == kofu and l["moku"] == mk]
            kv = [sum(d8["ws"].cell(r, c).value or 0 for c in cols)
                  for r in d8["ken"]]
            mok.append([f"{kofu} 目標{mk}", me, round(av, 1),
                        round(sum(kv) / len(kv), 1), rk, kr])
            print(f"     {kofu} 目標{mk}　{me:5.0f}点　全国平均{av:5.1f}　"
                  f"県平均{sum(kv) / len(kv):5.1f}　全国{rk:>5}位　県内{kr:>2}位")

    # ══════ 書き出し
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = sheet(wb, "00_この表について", ["項", "内容"], [24, 120])
    put(ws, [
        ["何の表か", "保険者機能強化推進交付金・介護保険保険者努力支援交付金の"
         "令和6〜8年度の本町の得点を、原典の全国一覧表から枝の単位で数え直し、"
         "推移と位置を整理したもの。"],
        ["満点", "⚠ 800点（推進400点・支援400点）。枝の配点を足すと929点に"
         "なるが、選択肢が排他の枝を含む見かけの合計である。失点は800−得点。"],
        ["成果指標群", "⚠ 目標Ⅳは推進・支援で同一の指標・同一の得点である。"
         "点は両交付金に別に入るので足してよいが、項目の数は1度だけ数える。"],
        ["推移の比べ方", "⚠ 得点率（得点÷その年度の配点）で比べる。素点で"
         "比べると、国が配点を下げただけの項目が「悪化」に入る。"],
        ["県内の比較", "宮城県内の保険者について、分布（平均・中央値・"
         "四分位・最大・最小）と本町の順位のみを示す。他団体の名称・個別の"
         "得点は書き出さない。"],
        ["交付額", "令和6年度の原典（交付見込額確定版）に交付見込額がある。"
         "令和7・8年度の原典には交付額の列がない。"],
    ])
    ws = sheet(wb, "01_合計の推移", ["年度", "本町", "全国平均", "全国順位",
               "保険者数", "宮城県平均", "宮城県 中央値", "県内順位",
               "県内保険者数", "全国平均との差"],
               [10, 8, 10, 10, 10, 11, 12, 10, 12, 14])
    put(ws, [[f"令和{y}年度", i["me"], round(i["avg"], 1), i["rank"], i["n"],
              round(i["kavg"], 1), i["kmed"], i["krank"], i["kn"],
              round(i["me"] - i["avg"], 1)] for y, i in ichi.items()])
    ws = sheet(wb, "02_低下の内訳_R7→R8",
               ["区分", "交付金", "目標", "群", "項目", "枝", "令和7 得点",
                "令和7 配点", "令和8 得点", "令和8 配点", "増減"],
               [24, 7, 6, 6, 44, 6, 10, 10, 10, 10, 8])
    r2 = []
    for b, lst in bunrui.items():
        for kk, a7, h7, a8, h8, d in sorted(lst, key=lambda x: x[5]):
            r2.append([b, kk[0], kk[1], kk[2], kk[3], kk[4],
                       a7 if a7 is not None else "―",
                       h7 if h7 is not None else "―",
                       a8 if a8 is not None else "―",
                       h8 if h8 is not None else "―", d])
    put(ws, r2, lambda r: WARN if r[0].startswith("取組の後退") else
        (OK if r[0].startswith("取組の前進") else None))
    ws = sheet(wb, "03_素点と得点率",
               ["交付金", "目標", "群", "項目", "令和6 得点", "令和6 配点",
                "令和8 得点", "令和8 配点", "判定"],
               [7, 6, 6, 50, 10, 10, 10, 10, 40])
    put(ws, [[ik[0], ik[1], ik[2], ik[3], a6, h6, a8, h8,
              "⚠ 素点では悪化。得点率は下がっていない（配点の改定）"]
             for ik, a6, h6, a8, h8 in haiten_dake], lambda r: WARN)
    ws = sheet(wb, "04_目標ごとの位置_R8",
               ["目標", "本町", "全国平均", "宮城県平均", "全国順位",
                "県内順位"], [14, 8, 10, 11, 10, 10])
    put(ws, mok)
    wb.save(OUT)
    print()
    print(f"  {OUT}")

    # ══════ 検査8・9
    print("  ── 検査8　⚠ 他団体の名称を書き出していないか")
    wb2 = openpyxl.load_workbook(OUT, data_only=True)
    zen = "\n".join(str(v) for sh in wb2 for row in sh.iter_rows(values_only=True)
                    for v in row if v is not None)
    ken_na = {s(Y[8]["ws"].cell(r, 8).value) for r in Y[8]["ken"]} - {MACHI}
    judge(not any(n and n in zen for n in ken_na),
          f"宮城県内の他の{len(ken_na)}保険者の名称が0件")
    print("  ── 検査9　満点を800点としている")
    judge("929" not in zen.replace("929点になる", ""),
          "見かけの合計929点を満点として使っていない")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜9　不適合0件")


if __name__ == "__main__":
    main()
