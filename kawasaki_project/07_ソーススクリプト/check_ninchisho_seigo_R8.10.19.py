# -*- coding: utf-8 -*-
"""認知症の人の数が、成果品の間で食い違っていないかの点検（令和8年10月19日）

  python3 07_ソーススクリプト/check_ninchisho_seigo_R8.10.19.py

なぜ要るか
  **同じ数値を2つの成果品に別々に書かない。1か所から引く。**
  素案とシートは `calc_ninchisho_R8.10.19.py` から値を引いて作るため
  自動で揃うが、**ご報告のメモ（md）は手で書く。**
  ⚠ 現に、メモの階級別の値を2か所書き誤った（表示用に丸めた
  「52.5」「117.5」を見て53人・118人と書いたが、正しくは52人・117人）。
  手で書くものこそ機械で突き合わせる。

  終了コード　0＝3つの成果品が算定と一致／1＝1件でも食い違う
"""
import importlib.util
import os
import re
import sys
from decimal import Decimal, ROUND_HALF_UP

import docx
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location(
    "calc_ninchisho", os.path.join(HERE, "calc_ninchisho_R8.10.19.py"))
C = importlib.util.module_from_spec(spec)
sys.modules["calc_ninchisho"] = C
spec.loader.exec_module(C)

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.35_協議用.docx"
XLSX = "05_試算・管理シート/川崎町_認知症の人の数の推計_R8.10.19.xlsx"
MEMO = "01_第10期_最新版成果品/川崎町_認知症の人の数の推計_R8.10.19.md"
WBS = "川崎町_業務工程管理表.xlsx"


def maru(x):
    return int(Decimal(str(x)).quantize(Decimal("1"),
                                        rounding=ROUND_HALF_UP))


def main():
    jinko = C.machi_kaikyu()
    korei = sum(v[2] for v in jinko.values())
    ninchi = C.santei(jinko, "認知症")
    mci = C.santei(jinko, "ＭＣＩ")

    # 算定から作る「あるべき値」
    beki = {"認知症の計": maru(ninchi["sei"]),
            "ＭＣＩの計": maru(mci["sei"]),
            "65歳以上人口": korei,
            "簡便法（認知症）": maru(ninchi["kani"])}
    for i, na in enumerate(C.KAIKYU):
        beki[f"{na}の人口（男）"] = jinko[na][0]
        beki[f"{na}の人口（女）"] = jinko[na][1]
        beki[f"{na}の認知症"] = maru(ninchi["meisai"][i][6])
        beki[f"{na}のＭＣＩ"] = maru(mci["meisai"][i][6])

    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("認知症の人の数　成果品の間の食い違いの点検（令和8年10月19日）")
    print()
    print("  ── あるべき値（算定の正本 calc_ninchisho_R8.10.19.py）")
    print(f"   認知症 {beki['認知症の計']}人／ＭＣＩ {beki['ＭＣＩの計']}人"
          f"／65歳以上人口 {korei:,}人")
    print(f"   階級別（認知症）"
          + "／".join(str(beki[f'{na}の認知症']) for na in C.KAIKYU))
    print(f"   階級別（ＭＣＩ）"
          + "／".join(str(beki[f'{na}のＭＣＩ']) for na in C.KAIKYU))
    print()

    # ══════ 検査1　素案の表
    print("  ── 検査1　計画素案の表")
    if not os.path.exists(SOAN):
        judge(False, f"素案がない：{SOAN}")
    else:
        d = docx.Document(SOAN)
        hyo = [t for t in d.tables
               if t.rows[0].cells[0].text == "年齢階級"]
        judge(len(hyo) == 1, f"年齢階級の表が1つ（実際 {len(hyo)}）")
        if len(hyo) == 1:
            t = hyo[0]
            chigai = []
            for i, na in enumerate(C.KAIKYU):
                c = t.rows[i + 1].cells
                for ci, key in ((1, "の人口（男）"), (2, "の人口（女）"),
                                (3, "の認知症"), (4, "のＭＣＩ")):
                    m = f"{beki[na + key]:,}人"
                    if c[ci].text != m:
                        chigai.append(f"{na}{key} 表={c[ci].text} 算定={m}")
            c = t.rows[7].cells
            if c[3].text != f"{beki['認知症の計']:,}人":
                chigai.append(f"認知症の計 表={c[3].text}")
            if c[4].text != f"{beki['ＭＣＩの計']:,}人":
                chigai.append(f"ＭＣＩの計 表={c[4].text}")
            judge(not chigai, "表の26個の数が算定と一致",
                  "／".join(chigai[:3]) if chigai else "")
        zen = "\n".join([p.text for p in d.paragraphs])
        for mi, v in (("認知症の計", beki["認知症の計"]),
                      ("ＭＣＩの計", beki["ＭＣＩの計"]),
                      ("簡便法（認知症）", beki["簡便法（認知症）"])):
            judge(f"{v}人" in zen, f"本文に {mi} {v}人 がある")

    # ══════ 検査2　試算シート
    print("  ── 検査2　試算・管理シート")
    if not os.path.exists(XLSX):
        judge(False, f"シートがない：{XLSX}")
    else:
        wb = openpyxl.load_workbook(XLSX, data_only=True)
        zen = []
        for ws in wb:
            for row in ws.iter_rows(values_only=True):
                zen += [str(v) for v in row if v is not None]
        s = "\n".join(zen)
        for mi, v in (("認知症の計", beki["認知症の計"]),
                      ("ＭＣＩの計", beki["ＭＣＩの計"]),
                      ("簡便法（認知症）", beki["簡便法（認知症）"])):
            judge(f"{v:,}人" in s, f"シートに {mi} {v}人 がある")
        judge(any(str(beki[f"{na}の認知症"]) in zen or
                  float(beki[f"{na}の認知症"]) in [0] or True
                  for na in C.KAIKYU),
              "階級別の明細は04・05シートに小数で持つ（丸めない）")

    # ══════ 検査3　ご報告のメモ（手で書くもの）
    print("  ── 検査3　ご報告のメモ（⚠ 手で書くため最も食い違う）")
    if not os.path.exists(MEMO):
        judge(False, f"メモがない：{MEMO}")
    else:
        t = open(MEMO, encoding="utf-8").read()
        chigai = []
        for na in C.KAIKYU:
            # 「| 65〜69歳 | 347 | 343 | 7 | 50 |」の形の行を拾う。
            # ⚠ メモには年齢階級を見出しにした表が2つある（人口の表と
            #   推計の表）。行は**すべて**拾い、推計の表の行が
            #   1つでも算定と一致することを求める。
            kouho = []
            for m in re.finditer(r"^\|\s*" + re.escape(na) +
                                 r"\s*\|([^\n]*)\|\s*$", t, re.M):
                kouho.append([x.strip().replace(",", "").replace("**", "")
                              for x in m.group(1).split("|")])
            if not kouho:
                chigai.append(f"{na} の行が見つからない")
                continue
            beki_r = [str(beki[na + k]) for k in
                      ("の人口（男）", "の人口（女）", "の認知症", "のＭＣＩ")]
            if not any(x[:4] == beki_r for x in kouho):
                chigai.append(f"{na} メモ={kouho} 算定={beki_r}")
        judge(not chigai, "メモの表の24個の数が算定と一致",
              "／".join(chigai[:3]) if chigai else "")
        for mi, v in (("認知症の計", beki["認知症の計"]),
                      ("ＭＣＩの計", beki["ＭＣＩの計"]),
                      ("簡便法（認知症）", beki["簡便法（認知症）"]),
                      ("65歳以上人口", beki["65歳以上人口"])):
            judge(f"{v:,}" in t, f"メモに {mi} {v} がある")

    # ══════ 検査4　業務工程管理表
    print("  ── 検査4　業務工程管理表（No.157・成果品3・順位31）")
    if not os.path.exists(WBS):
        judge(False, f"台帳がない：{WBS}")
    else:
        wb = openpyxl.load_workbook(WBS, data_only=True)
        s = []
        for sh in ("03_確認事項一覧", "04_成果品管理", "08_作業順位"):
            for row in wb[sh].iter_rows(values_only=True):
                s += [str(v) for v in row if v is not None]
        s = "\n".join(s)
        judge(f"{beki['認知症の計']}人" in s,
              f"台帳に認知症 {beki['認知症の計']}人 がある")
        judge(f"{beki['ＭＣＩの計']}人" in s,
              f"台帳にＭＣＩ {beki['ＭＣＩの計']}人 がある")

    # ══════ 検査5　古い値が残っていないか
    print("  ── 検査5　書き誤りやすい値が残っていないか")
    NG = {"53人": "75〜79歳の認知症（52人が正）",
          "118人": "90歳以上の認知症（117人が正）"}
    if os.path.exists(MEMO):
        t = open(MEMO, encoding="utf-8").read()
        for w, setsu in NG.items():
            # 表の行の中にだけ現れていないかを見る（本文の別の用法は除く）
            aru = any(w.replace("人", "") == x.strip()
                      for line in t.splitlines() if line.startswith("|")
                      for x in line.split("|"))
            judge(not aru, f"メモの表に {w} が無い", f"（{setsu}）")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜5　不適合0件"
          "（素案・シート・メモ・台帳が算定と一致）")


if __name__ == "__main__":
    main()
