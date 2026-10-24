# -*- coding: utf-8 -*-
"""第2回策定委員会 想定問答集 v7 → v8（町内の軽費老人ホームの廃止の確認結果）
（令和8年10月22日）

  python3 07_ソーススクリプト/build_2kai_qa_v8.py

v7 からの変更の理由
  計画素案を Ver.2.25、委員会資料を v12 に上げた。同じ版で直す。
  (1) QB-62 の A と根拠データ：廃止日・入所者の推移・退所先を書き、
      ⚠ 「住まい70人分が失われた」を改める（定員と入所者の別）
  (2) 計画素案の参照 Ver.2.24 → Ver.2.25
  問の数は変えない（全67問）。
"""
import importlib.util
import os
import re
import shutil
import sys

import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def load(name, fn):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(spec)
    sys.modules[name] = m
    spec.loader.exec_module(m)
    return m


S25 = load("fix_soan_v2250", "fix_soan_v2250_keihi.py")
A, GEN = S25.A, S25.GEN

SRC = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v7.docx"
DST = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v8.docx"
KEN = "宮城県「老人福祉施設等入所状況調」"

NAOSU = [
    ("A．宮城県「老人福祉施設等入所状況調」によると、町内にあった軽費老人ホーム",
     f"A．町内にあった軽費老人ホーム（定員{A['定員']}人）は、{S25.HAISHI}を"
     f"もって事業を廃止しました。{KEN}によると、入所者は令和5年4月1日時点で"
     f"{GEN['R5.4.1']}人、7月1日時点で{GEN['R5.7.1']}人と減り、10月1日時点で"
     f"{GEN['R5.10.1']}人となって休止中とされています。令和5年1月から9月まで"
     f"に退所した{A['退所']}人の退所先は、他の施設{A['他施設']}人、入院"
     f"{A['入院']}人、在宅{A['在宅']}人です。⚠ 定員の{A['定員']}人が住まいを"
     "失ったのではありません。なくなったのは、町内の高齢者向けの住まいの"
     f"選択肢として定員{A['定員']}人分です。⚠ 入所されていた方のうち本町の"
     "住民の数と、移られた先の施設の種別・市町村は、公表資料にはありません。"
     "廃止の理由と入所されていた方への措置は、社会福祉法第64条による宮城県"
     "への廃止届に記載される事項であり、その内容を確認しています。"
     "軽費老人ホームは老人福祉法に基づく施設で、介護保険の指定事業所では"
     "ないため、介護サービス情報公表システムには載りません。"),
    ("【根拠データ】 宮城県「老人福祉施設等入所状況調」（令和5年10月1日「休止中」",
     f"【根拠データ】 {KEN}（令和5年4月1日・7月1日・10月1日の入所現員と"
     "退所理由、令和6年1月1日「休止中」・4月1日以降掲載なし）。廃止日は"
     "法人の公表による。宮城県 軽費老人ホーム廃止届（様式第24号）の記載事項"),
]


def txt(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("第2回策定委員会 想定問答集 v7 → v8（軽費老人ホームの廃止の確認結果）")
    print()
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body
    from fix_soan_v111 import set_el

    n = 0
    for atama, atarashii in NAOSU:
        hit = [el for el in body.iterchildren()
               if el.tag == qn("w:p") and txt(el).startswith(atama)]
        if len(hit) != 1:
            judge(False, f"『{atama[:20]}…』に当たる段落が{len(hit)}件")
            continue
        set_el(hit[0], atarashii)
        n += 1
    judge(n == len(NAOSU), f"{len(NAOSU)}段落を書き改めた（実際 {n}）")

    n_ver = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(x.text or "" for x in el.iter(qn("w:t")))
        if "素案Ver.2.24" in s:
            set_el(el, s.replace("Ver.2.24", "Ver.2.25"))
            n_ver += 1
    judge(n_ver >= 30, f"素案の版の参照を {n_ver}か所改めた")
    doc.save(DST)

    d2 = docx.Document(DST)
    zen = "\n".join(p.text for p in d2.paragraphs) + "\n" + "\n".join(
        c.text for t in d2.tables for r in t.rows for c in r.cells)
    print("  ── 検査1　件数（変えていない）")
    qs = [int(x) for x in re.findall(r"QB-(\d+)　", zen)]
    judge(sorted(qs) == list(range(1, 68)), "全67問・欠番なし")
    print("  ── 検査2　内容")
    judge("住まい70人分が失われた" not in zen, "「住まい70人分が失われた」が0件")
    judge(S25.HAISHI in zen and f"他の施設{A['他施設']}人" in zen,
          "廃止日と退所先がある")
    judge("Ver.2.24" not in zen, "「Ver.2.24」が0件")
    print("  ── 検査3　⚠ 施設名・法人名・他団体名")
    for na in ("蔵王グリーンホーム", "鶴寿会", "0224-84-4853", "0224-84-2262",
               "釜房みどりの園", "アルパイン", "ながさか", "リリーハイム",
               "北塩原", "金ヶ崎", "川崎市"):
        judge(na not in zen, f"「{na}」が入っていない")
    judge("**" not in zen, "強調記号（**）が0件")

    print()
    print(f"  {DST}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 検査1〜3　不適合0件")


if __name__ == "__main__":
    main()
