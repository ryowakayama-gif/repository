# -*- coding: utf-8 -*-
"""第2回策定委員会 想定問答集 v8 → v9（資料 v13 に合わせる）
（令和8年10月22日）

  python3 07_ソーススクリプト/build_2kai_qa_v9.py

v8 からの変更の理由
  委員会資料を v13 に上げた（build_2kai_v13.py）。計画素案は策定委員会に
  未提出であるため修正履歴を資料に書かない、とのご指示により、資料から
  「当社資料の訂正（11件）」の節を除いた。問答集を同じ内容に揃える。
  (1) QB-45「訂正が11件に増えたが…」→ 資料にない前提の問になったため、
      「資料の数値は信用してよいのか」に置き換える（番号は変えない）
  (2) QB-40・QB-56 の頁数を資料 v13 に揃える
      ⚠ QB-40 は本文109頁・圧縮案102頁（Ver.2.9 の頃）のまま、
        QB-56 は本文117頁のままであった → 本文118頁・計131頁・圧縮後106頁程度
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

SRC = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v8.docx"
DST = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v9.docx"
SHIRYO = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v13.docx"
HEN = 118

NAOSU = [
    ("QB-45　【委員長】　訂正が11件に増えたが",
     "QB-45　【委員長】　資料の数値は信用してよいのか"),
    ("A．訂正の内容は、いずれも数値の読み方と出所の取り違えであり、",
     "A．保険料の試算は、算定式に印字されている値だけを用いて全39項目を"
     "検算しています。給付費は介護保険事業状況報告（年報・令和7年度）と、"
     "令和7年度の基準値は地域包括ケア「見える化」システムの総括表と32サービス"
     "すべてで突き合わせています。供給側の数値は介護サービス情報公表システムと"
     "宮城県の入所状況調という公表資料によっています。同じ数値を2つの資料に"
     "別々に書かず、出所を1か所に固定しています。"),
    ("【根拠データ】 第2回資料 ２（訂正の表・11件）。",
     "【根拠データ】 第2回資料 ６（第１０期の算定・保険料を動かす要因）。"
     "保険料の検算39項目すべて適合"),
    ("A．委託仕様書は「100頁程度」としています。現在の素案は本文109頁",
     f"A．委託仕様書は「100頁程度」としています。現在の素案は本文{HEN}頁・"
     f"資料編（見込13頁）の計{HEN + 13}頁です。詳細な分析（交付金の枝番の"
     "分析、保険料の感度分析、認知症施策の追加候補）を別冊に移すと、本文は"
     "106頁程度・計119頁程度となります。「100頁程度」を本文のみと読む場合でも、"
     "なお6頁程度の圧縮を要します。本文のみと読むか資料編を含めるかを、町と"
     "ご相談のうえ定めたいと考えています。"),
    ("【根拠データ】 第2回資料 ８（計画書の頁数の見通しの表）。素案Ver.2.9",
     f"【根拠データ】 第2回資料 ８（計画書の頁数の見通しの表）。素案 本文"
     f"{HEN}頁、別冊に移した場合 本文106頁程度"),
    ("QB-56　【住民代表】　計画書が117頁では、",
     f"QB-56　【住民代表】　計画書が{HEN}頁では、仕様書の「100頁程度」を"
     "超えるのではないか"),
    ("A．超えます。現在の素案は本文117頁で、",
     f"A．超えます。現在の素案は本文{HEN}頁で、資料編（見込13頁）を加えると"
     f"計{HEN + 13}頁です。詳細な分析を別冊に移すと本文は106頁程度になります。"
     "「100頁程度」を本文のみと読むか資料編を含めて読むかで扱いが変わります"
     "ので、町とご相談のうえ定めたいと考えています。"),
    ("【根拠データ】 第2回資料 ８（頁数の見通しの表）。素案Ver.2.25 本文117頁",
     f"【根拠データ】 第2回資料 ８（頁数の見通しの表）。素案Ver.2.25 本文{HEN}頁"),
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

    print("第2回策定委員会 想定問答集 v8 → v9（資料 v13 に合わせる）")
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
            judge(False, f"『{atama[:24]}…』に当たる段落が{len(hit)}件")
            continue
        set_el(hit[0], atarashii)
        n += 1
    judge(n == len(NAOSU), f"{len(NAOSU)}段落を書き改めた（実際 {n}）")
    doc.save(DST)

    d2 = docx.Document(DST)
    zen = "\n".join(p.text for p in d2.paragraphs) + "\n" + "\n".join(
        c.text for t in d2.tables for r in t.rows for c in r.cells)
    sh = docx.Document(SHIRYO)
    st = "\n".join([p.text for p in sh.paragraphs]
                   + [c.text for t in sh.tables for r in t.rows for c in r.cells])
    print("  ── 検査1　件数（変えていない）")
    qs = [int(x) for x in re.findall(r"QB-(\d+)　", zen)]
    judge(sorted(qs) == list(range(1, 68)), "全67問・欠番なし")
    print("  ── 検査2　資料 v13 と同じ内容")
    judge("訂正の表・11件" not in zen and "訂正が11件" not in zen,
          "資料から除いた訂正11件を前提にした記述が0件")
    judge("本資料の訂正" not in st, "資料 v13 にも訂正11件がない")
    judge("109頁" not in zen and "102頁" not in zen and "117頁" not in zen,
          "古い頁数（109・102・117頁）が0件")
    judge(f"本文{HEN}頁" in zen and f"{HEN}頁" in st and "106頁程度" in st,
          f"本文{HEN}頁・106頁程度（資料と同じ）")
    print("  ── 検査3　施設名・法人名・他団体名")
    for na in ("蔵王グリーンホーム", "鶴寿会", "釜房みどりの園", "アルパイン",
               "ながさか", "リリーハイム", "北塩原", "金ヶ崎", "川崎市"):
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
